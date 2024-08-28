import streamlit as st
import pandas as pd
from langchain_experimental.agents.agent_toolkits import create_pandas_dataframe_agent
from langchain.agents.agent_types import AgentType
from langchain.chat_models import ChatOpenAI
from langchain.memory import ConversationBufferMemory
from dotenv import load_dotenv
import os

from extractor import ExcelTableExtractor


def initialize_session_state():
    """Initialize session state variables."""
    if "messages" not in st.session_state:
        st.session_state["messages"] = []
    if "conversation_memory" not in st.session_state:
        st.session_state["conversation_memory"] = ConversationBufferMemory(
            memory_key="chat_history"
        )


def display_chat_history():
    """Display the chat history from the session state."""
    for message in st.session_state["messages"]:
        with st.chat_message(message.get("role")):
            st.write(message.get("content"))


def load_data(file_path):
    """Load data from an Excel file."""
    return pd.read_excel(file_path)


def get_user_input():
    """Get user input from the chat interface."""
    return st.chat_input("Say something")


def main():
    """Main function to run the Streamlit app."""
    load_dotenv()
    api_key = os.getenv("OPENAI_API_KEY")

    initialize_session_state()
    display_chat_history()

    st.title("Excel Data Chatbot")

    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
    if uploaded_file is not None:
        extractor = ExcelTableExtractor(uploaded_file)
        extractor.extract_tables()
        dfs = extractor.to_dataframes()

        prompt = get_user_input()

        if prompt:
            # Add user message to session state
            st.session_state["messages"].append({"role": "user", "content": prompt})

            # Display user message
            with st.chat_message("user"):
                st.write(prompt)

            # Initialize language model
            llm = ChatOpenAI(temperature=0.9, model="gpt-4o", openai_api_key=api_key)

            # Create agent
            agent = create_pandas_dataframe_agent(
                llm,
                dfs,
                handle_parsing_errors=True,
                verbose=True,
                agent_type=AgentType.OPENAI_FUNCTIONS,
                number_of_head_rows=10,
            )

            # Get response from agent
            response = agent.run(prompt)

            # Add assistant message to session state
            st.session_state["messages"].append(
                {"role": "assistant", "content": response}
            )

            # Display assistant message
            with st.chat_message("assistant"):
                st.write(response)


if __name__ == "__main__":
    main()

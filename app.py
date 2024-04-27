# __import__('pysqlite3')
# import sys
# sys.modules['sqlite3'] = sys.modules.pop('pysqlite3')
import ast

import pandas as pd
import streamlit as st
from dotenv import load_dotenv
from PyPDF2 import PdfReader
from langchain_core.messages import AIMessage, HumanMessage
from langchain_community.document_loaders import WebBaseLoader
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain_community.vectorstores import Chroma
from langchain_openai import OpenAIEmbeddings, ChatOpenAI
from langchain_core.prompts import ChatPromptTemplate, MessagesPlaceholder
from langchain.chains import create_history_aware_retriever, create_retrieval_chain
from langchain.chains.combine_documents import create_stuff_documents_chain
from htmlTemplates import css, bot_template, user_template


def get_vectorstore_from_text(text):
    # split the text into chunks
    text_splitter = RecursiveCharacterTextSplitter()
    text_chunks = text_splitter.split_text(text)

    # create a vectorstore from the chunks
    vector_store = Chroma.from_texts(text_chunks, OpenAIEmbeddings())

    return vector_store


def get_conversation_chain(vector_store):
    llm = ChatOpenAI()

    retriever = vector_store.as_retriever()

    prompt = ChatPromptTemplate.from_messages([
        MessagesPlaceholder(variable_name="chat_history"),
        ("user", "{input}"),
        ("user",
         "Given the above conversation, generate a search query to look up "
         "in order to get information relevant to the conversation")
    ])

    retriever_chain = create_history_aware_retriever(llm, retriever, prompt)

    prompt = ChatPromptTemplate.from_messages([
        ("system", "Answer the user's questions based on the below context:\n\n{context}"),
        MessagesPlaceholder(variable_name="chat_history"),
        ("user", "{input}"),
    ])

    stuff_documents_chain = create_stuff_documents_chain(llm, prompt)

    conversation_rag_chain = create_retrieval_chain(retriever_chain, stuff_documents_chain)

    return conversation_rag_chain


def get_response(user_input):
    # Check if 'conversation' has been initialized
    if st.session_state.conversation is not None:
        response = st.session_state.conversation.invoke({
            "chat_history": st.session_state.chat_history,
            "input": user_input
        })
        return response['answer']
    else:
        return "La conversation n'a pas encore été initialisée. Veuillez fournir une source de texte d'abord."


def get_pdf_text(pdf_docs):
    text = ""
    for pdf in pdf_docs:
        pdf_reader = PdfReader(pdf)
        for page in pdf_reader.pages:
            text += page.extract_text()
            # st.write(text)
    return text


def get_excel_text(excel_files):
    text = ""
    for excel in excel_files:
        excel_reader = pd.ExcelFile(excel)
        for sheet_name in excel_reader.sheet_names:
            df = excel_reader.parse(sheet_name)
            text += df.to_string(index=False, header=False)
    return text


def get_url_text(url):
    # get the text in document form
    loader = WebBaseLoader(url)
    document = loader.load()
    if document:
        # Supposons que le premier document contient le contenu de la page
        page_content = document[0].page_content
        # st.write(page_content)
        return page_content
    else:
        return "Impossible de charger le contenu de la page."


def handle_userinput(user_question):
    # invoke the conversation chain and get the response
    response = get_response(user_question)
    st.session_state.chat_history.append(HumanMessage(content=user_question))
    st.session_state.chat_history.append(AIMessage(content=response))

    # display the chat messages with streamlit and html templates
    for i, message in enumerate(st.session_state.chat_history):
        if i % 2 == 0:
            st.write(bot_template.replace("{{MSG}}", message.content), unsafe_allow_html=True)
        else:
            st.write(user_template.replace("{{MSG}}", message.content), unsafe_allow_html=True)


def main():
    load_dotenv()
    st.set_page_config(page_title="EPL Team Chatbot", page_icon="🎓")
    st.title("EPL Team Chatbot")
    st.write(css, unsafe_allow_html=True)

    if "conversation" not in st.session_state:
        st.session_state.conversation = None
    if "chat_history" not in st.session_state:
        st.session_state.chat_history = []
    if "source_type" not in st.session_state:
        st.session_state.source_type = None

    st.header("Chat with a URL, a PDF, or an Excel document :books:")
    user_question = st.chat_input("Ask a question about your source:")
    if user_question:
        handle_userinput(user_question)

    with st.sidebar:
        st.header("EPL Team Resources")
        source_type = st.radio("Choose your source type:", ("URL", "PDF", "Excel"))
        st.session_state.source_type = source_type
        if source_type == "URL":
            url = st.text_input("Enter your URL here and click on 'Process'")
            if st.button("Process"):
                with st.spinner("Processing"):
                    text = get_url_text(url)
                    vectorstore = get_vectorstore_from_text(text)
                    st.session_state.conversation = get_conversation_chain(vectorstore)
                    st.session_state.chat_history = [
                        AIMessage(content="Hello, I am the EPL team URL Chatbot. How may I assist you today?"),
                    ]
        elif source_type == "PDF":
            pdf_docs = st.file_uploader("Upload your PDFs here and click on 'Process'", accept_multiple_files=True)
            if st.button("Process"):
                with st.spinner("Processing"):
                    text = get_pdf_text(pdf_docs)
                    vectorstore = get_vectorstore_from_text(text)
                    st.session_state.conversation = get_conversation_chain(vectorstore)
                    st.session_state.chat_history = [
                        AIMessage(content="Hello, I am the EPL team PDF Chatbot. How may I assist you today?"),
                    ]
        elif source_type == "Excel":  # Add an Excel section
            excel_files = st.file_uploader("Upload your Excel files here and click on 'Process'",
                                           accept_multiple_files=True)
            if st.button("Process"):
                with st.spinner("Processing"):
                    text = get_excel_text(excel_files)
                    vectorstore = get_vectorstore_from_text(text)
                    st.session_state.conversation = get_conversation_chain(vectorstore)
                    st.session_state.chat_history = [
                        AIMessage(content="Hello, I am the EPL team Excel Chatbot. How may I assist you today?"),
                    ]


if __name__ == '__main__':
    main()

import openai
import streamlit as st
import json
import PyPDF2
import docx
import openpyxl
import fitz  # PyMuPDF
import pandas as pd
from langchain_openai import AzureChatOpenAI
from azure.search.documents import SearchClient
from azure.core.credentials import AzureKeyCredential

OPENAI_API_KEY = 'AloI7eJioMWY1wsd3nDgpsv6gYv3rlZfM92lXxIcUpDAKMj25tbCJQQJ99AKACYeBjFXJ3w3AAABACOGqcXh'  # Replace with your Azure API key'  # Replace with your Azure API key
AZURE_OPENAI_ENDPOINT = 'https://kareenaopenai2.openai.azure.com/'  # Replace with your Azure endpoint
OPENAI_DEPLOYMENT_NAME = 'TestingChatbots'  # The deployment name for the model you're using
OPENAI_API_VERSION = '2024-10-21'  # Azure OpenAI API version

# Initialize the AzureChatOpenAI model
llm = AzureChatOpenAI(
    azure_deployment=OPENAI_DEPLOYMENT_NAME,
    azure_endpoint=f"{AZURE_OPENAI_ENDPOINT}/openai/deployments/{OPENAI_DEPLOYMENT_NAME}/chat/completions?api-version={OPENAI_API_VERSION}",
    openai_api_key=OPENAI_API_KEY,
    openai_api_version=OPENAI_API_VERSION
)

# Azure Cognitive Search setup
search_service_name = "dgivazuresearch"
index_name = "sharepoint-index-meghan"
search_api_key = "CtiqITkpp2hlCNbLC4L9RpAXexFwb1a5S99waKbu0aAzSeDZkD6h"
endpoint = "https://dgivazuresearch.search.windows.net/"
search_client = SearchClient(endpoint=endpoint, index_name=index_name, credential=AzureKeyCredential(search_api_key))

def extract_text_from_pdf(file):
    """Extract text from an uploaded PDF file using PyMuPDF (better accuracy)."""
    text = ""
    pdf_doc = fitz.open(stream=file.read(), filetype="pdf")
    for page in pdf_doc:
        text += page.get_text("text") + "\n"
    return text.strip()

def extract_text_from_docx(file):
    """Extract text from an uploaded DOCX file."""
    doc = docx.Document(file)
    text = "\n".join([p.text for p in doc.paragraphs])
    return text

def extract_text_from_excel(file):
    """Extract text from an uploaded Excel file (.xlsx or .xls)."""
    try:
        df_dict = pd.read_excel(file, sheet_name=None)  # Read all sheets
        text = ""
        for sheet_name, df in df_dict.items():
            text += f"\n--- Sheet: {sheet_name} ---\n"
            text += df.to_string(index=False)  # Convert DataFrame to text
        return text.strip()
    except Exception as e:
        st.error(f"⚠️ Error reading Excel file: {e}")
        return None
    
def process_uploaded_file(file):
    """Process uploaded file and return extracted text."""
    if file.name.endswith(".pdf"):
        return extract_text_from_pdf(file)
    elif file.name.endswith(".docx"):
        return extract_text_from_docx(file)
    elif file.name.endswith((".xlsx", ".xls")):
        return extract_text_from_excel(file)
    else:
        st.error("⚠️ Unsupported file format. Only PDF, DOCX, and Excel (XLSX, XLS) are allowed.")
        return None

# Function to interact with the Azure OpenAI API
def get_openai_response(prompt):
    try:
        # Query the model with the user input
        response = llm.predict(prompt)
        return response.strip()  # Return the response after cleaning it up
    except Exception as e:
        return f"Error: {str(e)}"

# Function to query Azure Cognitive Search and get relevant content (limit results)
    
def search_index(query):
    """Search Azure Index for MEWAH-related content using Semantic Search."""
    results = search_client.search(
        search_text=query,
        top=3,
        query_type="semantic",
        semantic_configuration_name="meghansemantic"
    )

    relevance_threshold = 2.5  # Adjust for better accuracy

    filtered_docs = [doc['content'] for doc in results if doc.get('@search.score', 0) >= relevance_threshold]

    if filtered_docs:
        combined_content = "\n\n".join(filtered_docs)  
        summary_prompt = f"Summarize and answer strictly about Mewah:\n\nQuery: {query}\n\nContent:\n{combined_content}"
        return get_openai_response(summary_prompt)
    
    return "Sorry, this is not available in Mewah's indexed content."

# Custom CSS for styling the UI with chatbot-like design
# Page config
st.set_page_config(page_title="Azure Chatbot", layout="wide")
st.logo("https://iportal.ncheo.com/images/logo-Icon1.png") 

# Custom CSS for UI Styling
st.markdown("""
    <style>
        /* Global Styles */
        body {
            background-color: #f4f6f8;  
            font-family: 'Arial', sans-serif;
        }

        /* Button Styles */
        .stButton>button {
            background-color: #0078d4;
            color: white;
            font-weight: bold;
            border-radius: 10px;
            border: none;
            padding: 10px 20px;
            transition: background-color 0.3s ease;
        }
        .stButton>button:hover {
            background-color: #005a9e;
        }

        /* Custom File Upload Button */
        .file-upload-button {
            background-color: #0078d4;
            color: white;
            font-weight: bold;
            border-radius: 10px;
            border: none;
            padding: 8px 16px;
            cursor: pointer;
            display: inline-block;
            text-align: center;
            transition: background-color 0.3s ease;
        }
        .file-upload-button:hover {
            background-color: #005a9e;
        }
        .file-upload-button:active {
            background-color: #003d6b;
        }
        input[type="file"] {
            display: none;
        }

        /* Input and Text Area Styling */
        .stTextInput>div>input, .stTextArea>div>textarea {
            border-radius: 10px;
            border: 2px solid #0078d4;
            padding: 12px;
            width: 100%;
            background-color: white;
            color: #333333;
            transition: border 0.3s ease;
        }
        .stTextInput>div>input:focus, .stTextArea>div>textarea:focus {
            border: 2px solid #005a9e;
        }

        /* Radio Button Styling */
        .stRadio>div>label {
            font-size: 1.2em;
            font-weight: bold;
            color: #0078d4;
        }

        /* Chatbot Messages */
        .message-container {
            margin: 20px;
            padding: 10px;
            background-color: white;
            border-radius: 10px;
            box-shadow: 0px 4px 6px rgba(0, 0, 0, 0.1);
            overflow: hidden;
            max-height: 500px;
            overflow-y: auto;
        }
        .user-message {
            background-color: #0078d4;
            color: white;
            border-radius: 20px;
            padding: 12px 20px;
            margin: 10px 0;
            max-width: 70%;
            float: right;
            clear: both;
        }
        .bot-message {
            background-color: #e5e5e5;
            color: #333333;
            border-radius: 20px;
            padding: 12px 20px;
            margin: 10px 0;
            max-width: 70%;
            float: left;
            clear: both;
        }
        /* Sidebar always visible, not collapsible */
        .css-1d391kg {  /* This is the sidebar container class */
            position: relative !important;
            width: 250px !important;
            z-index: 2 !important;
        }

        .css-1d391kg .css-1v0mbdj {  /* Sidebar content styling */
            width: 100% !important;
        }

        /* Layout for Left-side elements */
        .left-container {
            width: 250px;
            padding: 20px;
            float: left;
        }

        /* Layout for Chat Area */
        .chat-area {
            margin-left: 280px;
            padding: 20px;
        }
    </style>
""", unsafe_allow_html=True)

# Ensure necessary session state variables exist
if "prev_query_type" not in st.session_state:
    st.session_state.prev_query_type = []
if "chat_history" not in st.session_state:
    st.session_state.chat_history = []
if "uploaded_files" not in st.session_state:
    st.session_state.uploaded_files = []

# Streamlit user interface
logo_path = "https://iportal.ncheo.com/images/logo-Icon1.png"  # Change this to the path of your logo file

# Custom HTML & CSS for perfect alignment of the title
st.markdown(
    f"""
    <div style="display: flex; align-items: center; justify-content: center; gap: 10px;">
        <img src="{logo_path}" width="50" style="flex-shrink: 0;">
        <h1 style="margin: 0; padding: 0; text-align: center; flex-grow: 1;">MEWAH AI CHATBOT</h1>
    </div>
    <p style="text-align: center; font-size: 16px; color: #555;">
        Your intelligent assistant for answering Mewah-specific and general queries.
    </p>
    """,
    unsafe_allow_html=True
)

# Sidebar for query type selection
st.sidebar.markdown("### Choose a query type:")
query_type = st.sidebar.radio("What would you like to ask?", ("Mewah-specific Query", "General Query"))

# Force reload if query type changes
if st.session_state.prev_query_type is not None and query_type != st.session_state.prev_query_type:
    st.session_state.chat_history = []  # Clear chat history
    st.session_state.uploaded_files = []  # Clear uploaded files
    st.session_state.prev_query_type = query_type  # Update the previous type
    st.rerun()  # Force app to refresh

# File uploader for "General Query", remains at the top
if query_type == "General Query":
    uploaded_files = st.file_uploader("Upload a file", type=["docx", "pdf", "xlsx", "xls"], accept_multiple_files=True)
    if uploaded_files:
        st.session_state.uploaded_files = uploaded_files

# Chat input for user messages
user_input = st.chat_input("Ask a question...")

# **Ensure conversation_history is initialized**
conversation_history = "\n".join([f"{role}: {message}" for role, message in st.session_state.chat_history])

if user_input:
    # Add user message to history
    st.session_state.chat_history.append(("User", user_input))

    with st.spinner('Thinking...'):
        # Update conversation history after user input
        conversation_history = "\n".join([f"{role}: {message}" for role, message in st.session_state.chat_history])

if user_input:
    if query_type == "Mewah-specific Query":
        response = search_index(user_input)
    elif query_type == "General Query":
        if st.session_state.uploaded_files:
            all_extracted_texts = ""
            for uploaded_file in st.session_state.uploaded_files:
                extracted_text = process_uploaded_file(uploaded_file)
                if extracted_text:
                    all_extracted_texts += f"\n\n### File: {uploaded_file.name} ###\n{extracted_text}"

            if all_extracted_texts:
                prompt = f"Documents content:\n\n{all_extracted_texts}\n\n{conversation_history}\n\nBot:"
                response = get_openai_response(prompt)
            else:
                response = "No content extracted from the uploaded files."
        else:
            prompt = f"{conversation_history}\n\nBot:"
            response = get_openai_response(prompt)

    # Add bot response to chat history
    st.session_state.chat_history.append(("Bot", response))

# Display chat history for back-to-back conversation
for role, message in st.session_state.chat_history:
    if role == "User":
        st.markdown(f'<div class="message-container"><div class="user-message">{message}</div></div>', unsafe_allow_html=True)
    else:
        st.markdown(f'<div class="message-container"><div class="bot-message">{message}</div></div>', unsafe_allow_html=True)

# Clear chat button
if st.button("Clear Chat"):
    st.session_state.chat_history = []  # Clear chat history
    st.session_state.uploaded_files = []  # Clear uploaded files
    st.rerun()  # Force Streamlit to refresh the app
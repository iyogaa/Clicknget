import streamlit as st

def apply_custom_css():
    st.markdown("""
    <style>
    .stApp {
        background-color: #000000;
    }
    [data-testid="stSidebar"] {
        background-color: #111111 !important;
    }
    /* Sidebar text bright white but now in normal font */
    [data-testid="stSidebar"] * {
        color: #FFFFFF !important;
        font-weight: normal;
    }
    /* Left-aligned Heading with smaller font and no border */
    .custom-heading {
        font-size: 2rem;
        color: white;
        text-align: left;
        font-weight: bold;
        margin-bottom: 1.5rem;
        margin-left: 2rem;
        background: none;
        border: none;
        padding: 0;
    }
    /* Remove extra empty box inside file uploader */
    [data-testid="stFileUploader"] > div {
        background-color: transparent !important;
        padding: 0 !important;
        margin: 0 !important;
        border: none !important;
        min-height: 0 !important;
        min-width: 0 !important;
    }
    /* Label and input text white */
    label, .stFileUploader, .stNumberInput label, .stSelectbox label {
        color: white !important;
    }
    /* White text for all content */
    body, .stMarkdown, .stText, .stDataFrame, .stMetric {
        color: white !important;
    }
    /* Custom button styling */
    .stButton>button {
        background-color: #000000;
        color: white;
        border-radius: 5px;
        padding: 0.5rem 1rem;
        font-weight: bold;
        border: 1px solid #333;
    }
    .stButton>button:hover {
        border-color: #4CAF50;
        color: #4CAF50;
    }
    /* Status indicator */
    .status-indicator {
        display: inline-block;
        width: 10px;
        height: 10px;
        border-radius: 50%;
        margin-right: 5px;
    }
    .status-operational {
        background-color: #4CAF50;
    }
    </style>
    """, unsafe_allow_html=True)

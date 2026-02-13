import streamlit as st
from accident_gpt.gpt import run

def run_accident_gpt():
    st.set_page_config(
    page_title="Accident GPT",
    page_icon="🚓",
    layout="wide",
    initial_sidebar_state="expanded"
    )
    st.header("Accident GPT 🚓")
    st.write("Use this tool to categorize incident descriptions into standardised accident categories.")
    st.info("""
    1. Use for Commercial Auto Clients.
    2. Upload only xlsx format file with lossrun_data sheet.
    3. Logic will work only on selected column.
    4. In output, AccidentCategory column will be added. Download as csv or excel.
    """)
    run()

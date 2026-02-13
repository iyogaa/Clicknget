import streamlit as st
from body_gpt.gpt import run

def run_body_gpt():
    st.header("Body GPT 🚓")
    st.write("Use this tool to categorize incident descriptions into standardised body hierarchy categories.")
    st.info("""
    1. Use for Worker Compensation Clients.
    2. Upload only xlsx format file with lossrun_data sheet.
    3. Logic will work only on selected column.
    4. In output, Body - Hierarchy 1 column will be added. Download as csv or excel.
    """)
    run()

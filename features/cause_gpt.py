import streamlit as st
from cause_gpt.gpt import run

def run_cause_gpt():
    st.header("Cause GPT 🚓")
    st.write("Use this tool to categorize incident descriptions into standardised cause categories.")
    st.info("""
    1. Use for Worker Compensation and General Liability Clients.
    2. Upload only xlsx format file with lossrun_data sheet.
    3. Logic will work only on selected column.
    4. In output, Cause - Standardised column will be added. Download as csv or excel.
    """)
    run()

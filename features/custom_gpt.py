import streamlit as st
from custom_gpt_categorisation.gpt import run

def run_custom_gpt():
    st.header("Custom Categorisation GPT 🚓")
    st.write("Use this tool to categorize incident descriptions into standardised injury type categories.")
    st.info("""
    1. Only for FCCI client.
    2. Upload only xlsx format file with lossrun_data sheet.
    3. Logic will work only on selected column.
    4. In output, Injury Type column will be added. Download as csv or excel.
    """)
    run()

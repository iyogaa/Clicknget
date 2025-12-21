import pandas as pd
import streamlit as st
import io
import os
import tempfile
from Pdf_maker import process_pdf

st.set_page_config(layout="wide")

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

def authenticate(username, password):
    if "credentials" in st.secrets and username in st.secrets["credentials"]:
        user_data = st.secrets["credentials"][username]
        if password == user_data["password"]:
            return user_data["role"]
    return None

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
    st.session_state["username"] = None
    st.session_state["role"] = None

def show_login():
    with st.sidebar:
        st.title("Login")
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        if st.button("Login"):
            role = authenticate(username, password)
            if role:
                st.session_state["authenticated"] = True
                st.session_state["username"] = username
                st.session_state["role"] = role
                st.rerun()
            else:
                st.error("Invalid username or password")

if not st.session_state["authenticated"]:
    show_login()
    st.stop()

def get_menu_options(role):
    base = ["MVR All Trans", "PDF Maker", "PDF Merger"]
    if role == "ADMIN":
        return base 
    elif role == "QA":
        return base
    elif role == "MAKER":
        return []
    return []

with st.sidebar:
    st.markdown(f"### 👋 Welcome, **{st.session_state['username']}**")
    st.markdown(f"**Role:** {st.session_state['role']}")
    st.markdown("---")

    menu_options = get_menu_options(st.session_state["role"])
    if menu_options:
        menu = st.radio("📋 Menu", menu_options, label_visibility="collapsed")
    else:
        st.warning("No menu options available.")
        menu = None

    st.markdown("---")
    if st.button("Logout"):
        st.session_state.clear()
        st.rerun()
    st.caption("Built with Yogaraj ")

if menu == "MVR All Trans":
    from Alltran import Alltrans
    
    st.title("Alltrans Process")

    main_file = st.file_uploader("Upload MVR File", type=["xlsx"])
    lookup_file = st.file_uploader("Upload Client Excel", type=["xlsx","CSV"])

    if main_file and lookup_file:
        if st.button("Process"):
            try:
                main_bytes = main_file.read()
                lookup_bytes = lookup_file.read()

            
                from openpyxl import load_workbook
                lookup_wb = load_workbook(io.BytesIO(lookup_bytes), read_only=True, data_only=True)
                sheets = lookup_wb.sheetnames
                chosen_sheet = None
                if len(sheets) > 1:
                    chosen_sheet = st.selectbox("Lookup workbook has multiple sheets. Choose one:", options=sheets)
                else:
                    chosen_sheet = sheets[0]
                    st.write(f"From Client Excel: {chosen_sheet}")

                gen = Alltrans(template_path="Template.xlsx",
                            alltrans_sheet="All Trans",
                            alltrans_header_row=4,#Fixed constand row for Column Headers
                            mvr_sheet_name="MVR")
                out = gen.run(main_bytes, lookup_bytes, chosen_lookup_sheet=chosen_sheet, preview_rows=8)

                original_name = getattr(main_file, "name", None)
                base = original_name.rsplit(".", 1)[0] if original_name else "Final_Report"
                out_name = f"{base}.xlsx"
                st.success("Final report generated")
                st.download_button("Download Final Report", data=out, file_name=out_name,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"Error: {e}")


elif menu == "PDF Maker":
    uploaded = st.file_uploader("Upload PDF", type=["pdf"])

    if uploaded:
        temp_input = "uploaded.pdf"
        with open(temp_input, "wb") as f:
            f.write(uploaded.read())

        if st.button("Process PDF"):
            st.info("Processing... bruhh...")

            output_file = "flattened.pdf"
            process_pdf(temp_input, output_file)

            with open(output_file, "rb") as f:
                st.download_button(
                    "Download Processed PDF",
                    f,
                    file_name="flattened.pdf",
                    mime="application/pdf"
                )

            # cleanup
            os.remove(temp_input)
            os.remove(output_file)

elif menu == "PDF Merger":
    from pdf_merge import merge_pdfs
    st.title("PDF Merger")

    uploaded_files = st.file_uploader(
        "Upload PDF files",
        type=["pdf"],
        accept_multiple_files=True
    )

    if uploaded_files and len(uploaded_files) >= 2:
        if st.button("Merge PDFs"):
            with st.spinner("Merging PDFs..."):
                merged_path = merge_pdfs(uploaded_files)

            with open(merged_path, "rb") as f:
                pdf_data = f.read()

            st.download_button(
                label="⬇ Download merged PDF",
                data=pdf_data,
                file_name="merged.pdf",
                mime="application/pdf"
            )
            
            os.remove(merged_path)

    elif uploaded_files:
        st.warning("Please upload at least 2 PDF files")
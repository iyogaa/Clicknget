import streamlit as st
import pandas as pd
import io
import os
import zipfile
from datetime import datetime
from PIL import Image
import fitz # PyMuPDF

from src.config.settings import config
from src.core.converter import ExcelToPdfConverter
from src.core.validators import validate_file
from src.core.exceptions import AppError
from src.utils.file_handler import temporary_file, temporary_directory
from src.utils.logger import logger

# --- PAGE CONFIG ---
st.set_page_config(
    page_title=config.APP_NAME,
    page_icon="📄",
    layout="wide",
)

# --- CSS ---
st.markdown("""
<style>
    .stButton>button {
        border-radius: 8px;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(0, 212, 170, 0.4);
    }
    .file-card {
        padding: 1rem;
        border-radius: 10px;
        background: #1e2129;
        margin-bottom: 0.5rem;
    }
</style>
""", unsafe_allow_html=True)

# --- SESSION STATE ---
if "history" not in st.session_state:
    st.session_state.history = []

# --- HELPERS ---
def get_pdf_preview(pdf_bytes: bytes):
    """Generate a preview image of the first page of a PDF."""
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        if len(doc) > 0:
            page = doc[0]
            pix = page.get_pixmap(matrix=fitz.Matrix(0.5, 0.5))
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            return img
    except Exception as e:
        logger.error(f"Preview failed: {e}")
    return None

# --- MAIN UI ---
st.title(f"🚀 {config.APP_NAME}")
st.write("Convert your Excel sheets into production-grade PDFs with ease.")

# Sidebar Settings
with st.sidebar:
    st.header("⚙️ Global Settings")
    page_size_name = st.selectbox("Page Size", ["A4", "Letter", "Legal"], index=0)
    orientation = st.selectbox("Orientation", ["Auto", "Portrait", "Landscape"], index=0)
    scale_mode = st.selectbox("Scale Mode", ["Fit to Width", "Actual Size"], index=0)
    
    from reportlab.lib.pagesizes import A4, LETTER, LEGAL
    page_sizes = {"A4": A4, "Letter": LETTER, "Legal": LEGAL}
    selected_page_size = page_sizes[page_size_name]
    if orientation == "Landscape":
        from reportlab.lib.pagesizes import landscape
        selected_page_size = landscape(selected_page_size)

    st.divider()
    st.header("🕒 History")
    if not st.session_state.history:
        st.info("No recent conversions.")
    else:
        for item in reversed(st.session_state.history):
            st.write(f"- {item['filename']} ({item['time']})")
        if st.button("Clear History"):
            st.session_state.history = []
            st.rerun()

# Layout
col_upload, col_process = st.columns([1, 1])

with col_upload:
    st.header("📤 Upload Files")
    uploaded_files = st.file_uploader(
        "Drop Excel files here (.xlsx, .xls)",
        type=list(config.ALLOWED_EXTENSIONS),
        accept_multiple_files=True,
        help="Max 50MB per file"
    )

with col_process:
    st.header("🛠️ Processing")
    if not uploaded_files:
        st.info("Upload files to start conversion.")
    else:
        if st.button("Convert All to PDF", type="primary"):
            results = []
            progress_bar = st.progress(0)
            
            for i, uploaded_file in enumerate(uploaded_files):
                try:
                    # Validate
                    validate_file(
                        uploaded_file.name, 
                        uploaded_file.size, 
                        config.ALLOWED_EXTENSIONS, 
                        config.MAX_UPLOAD_SIZE_MB
                    )
                    
                    with st.spinner(f"Converting {uploaded_file.name}..."):
                        converter = ExcelToPdfConverter(dpi=config.DEFAULT_DPI)
                        with temporary_file(suffix=".pdf") as out_path:
                            input_bytes = uploaded_file.getvalue()
                            converter.convert(
                                input_bytes, 
                                out_path, 
                                {"page_size": selected_page_size, "scale_mode": scale_mode}
                            )
                            
                            with open(out_path, "rb") as f:
                                pdf_bytes = f.read()
                            
                            results.append({
                                "name": f"{os.path.splitext(uploaded_file.name)[0]}.pdf",
                                "bytes": pdf_bytes
                            })
                            
                except AppError as ae:
                    st.error(f"❌ {uploaded_file.name}: {ae}")
                except Exception as e:
                    logger.error(f"Unexpected error: {e}")
                    st.error(f"❌ {uploaded_file.name}: Unexpected technical error.")
                
                progress_bar.progress((i + 1) / len(uploaded_files))
            
            # Finalize Batch
            if results:
                st.success(f"Converted {len(results)} files!")
                
                # Update history
                for res in results:
                    st.session_state.history.append({
                        "filename": res["name"],
                        "time": datetime.now().strftime("%H:%M:%S")
                    })
                if len(st.session_state.history) > config.HISTORY_LIMIT:
                    st.session_state.history = st.session_state.history[-config.HISTORY_LIMIT:]
                
                if len(results) == 1:
                    res = results[0]
                    st.download_button(
                        f"Download {res['name']}",
                        res['bytes'],
                        file_name=res['name'],
                        mime="application/pdf"
                    )
                    # Preview
                    preview_img = get_pdf_preview(res['bytes'])
                    if preview_img:
                        st.image(preview_img, caption="First Page Preview", use_container_width=True)
                else:
                    # Multi-file ZIP
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w") as zf:
                        for res in results:
                            zf.writestr(res["name"], res["bytes"])
                    
                    st.download_button(
                        "Download All as ZIP",
                        zip_buffer.getvalue(),
                        file_name="converted_pdfs.zip",
                        mime="application/zip"
                    )
                
                st.toast("Conversion complete!", icon="✅")

st.divider()
st.caption(f"© {datetime.now().year} {config.APP_NAME} | Optimized for Streamlit Cloud")

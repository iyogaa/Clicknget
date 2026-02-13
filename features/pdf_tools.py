import streamlit as st
import os
import tempfile
import io
import fitz
from Pdf_maker import process_pdf
from pdf_play import WordToPDF, ExcelToPDF, ImageToPDF, PDFCompressor, PDFMerger, PDFSplitter

def save_uploaded(uploaded):
    suffix = os.path.splitext(uploaded.name)[1]
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(uploaded.getbuffer())
        return tmp.name

def run_pdf_maker():
    st.title("PDF Flattening Tool")
    uploaded = st.file_uploader("Upload PDF", type=["pdf"])

    if uploaded:
        temp_input = "uploaded.pdf"
        with open(temp_input, "wb") as f:
            f.write(uploaded.read())

        if st.button("Process PDF"):
            st.info("Processing PDF...")

            output_file = "flattened.pdf"
            process_pdf(temp_input, output_file)

            with open(output_file, "rb") as f:
                st.download_button(
                    "Download Processed PDF",
                    f,
                    file_name="flattened.pdf",
                    mime="application/pdf"
                )

            os.remove(temp_input)
            os.remove(output_file)

def run_pdf_play():
    st.title("PDF Play Ground")

    tool_option = st.selectbox("Select Tool", [
        "Word → PDF", 
        "Excel → PDF", 
        "Images → PDF", 
        "Merge PDFs",
        "Split PDF",
        "Compress PDF"
    ])

    if tool_option == "Word → PDF":
        st.subheader("Word (.doc/.docx) to PDF")
        uploaded = st.file_uploader("Upload Word Document", type=["doc", "docx"])
        if uploaded and st.button("Convert"):
            with st.spinner("Converting Word to PDF..."):
                tmp_path = save_uploaded(uploaded)
                try:
                    converter = WordToPDF()
                    out_path = converter.convert(tmp_path)
                    with open(out_path, "rb") as f:
                        st.download_button("Download PDF", f, file_name=f"{uploaded.name}.pdf", mime="application/pdf")
                    os.remove(out_path)
                except Exception as e:
                    st.error(f"Error: {e}")
                finally:
                    os.remove(tmp_path)

    elif tool_option == "Excel → PDF":
        st.subheader("Excel (.xls/.xlsx) to PDF")
        uploaded = st.file_uploader("Upload Excel File", type=["xls", "xlsx"])
        if uploaded and st.button("Convert"):
            with st.spinner("Converting Excel to PDF..."):
                tmp_path = save_uploaded(uploaded)
                try:
                    converter = ExcelToPDF()
                    out_path = converter.convert(tmp_path)
                    with open(out_path, "rb") as f:
                        st.download_button("Download PDF", f, file_name=f"{uploaded.name}.pdf", mime="application/pdf")
                    os.remove(out_path)
                except Exception as e:
                    st.error(f"Error: {e}")
                finally:
                    os.remove(tmp_path)

    elif tool_option == "Images → PDF":
        st.subheader("Images to PDF")
        uploaded_files = st.file_uploader("Upload Images", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
        if uploaded_files and st.button("Convert"):
            with st.spinner("Converting Images to PDF..."):
                img_paths = []
                for up in uploaded_files:
                    img_paths.append(save_uploaded(up))
                try:
                    converter = ImageToPDF()
                    out_path = converter.convert(img_paths)
                    with open(out_path, "rb") as f:
                        st.download_button("Download PDF", f, file_name="images.pdf", mime="application/pdf")
                    if os.path.exists(out_path):
                        os.remove(out_path)
                except Exception as e:
                    st.error(f"Error: {e}")
                finally:
                    for p in img_paths:
                        if os.path.exists(p):
                            os.remove(p)

    elif tool_option == "Compress PDF":
        st.subheader("Compress PDF File Size")
        uploaded = st.file_uploader("Upload PDF", type=["pdf"])
        level = st.select_slider("Compression Level", options=["Low", "Medium", "High"], value="Medium")
        if uploaded and st.button("Compress"):
            with st.spinner("Compressing PDF..."):
                tmp_path = save_uploaded(uploaded)
                try:
                    compressor = PDFCompressor()
                    out_path = compressor.compress(tmp_path, level.lower())
                    original_size = os.path.getsize(tmp_path)
                    new_size = os.path.getsize(out_path)
                    reduction = (1 - new_size/original_size) * 100
                    st.success(f"Reduced by {reduction:.1f}% ({original_size/(1024*1024):.2f}MB → {new_size/(1024*1024):.2f}MB)")
                    with open(out_path, "rb") as f:
                        st.download_button("Download Compressed PDF", f, file_name=f"compressed_{uploaded.name}", mime="application/pdf")
                    os.remove(out_path)
                except Exception as e:
                    st.error(f"Error: {e}")
                finally:
                    os.remove(tmp_path)

    elif tool_option == "Merge PDFs":
        st.subheader("Merge Multiple PDFs")
        uploaded_files = st.file_uploader("Upload PDF files", type=["pdf"], accept_multiple_files=True)
        if "merge_order" not in st.session_state:
            st.session_state.merge_order = []
        def get_file_id(file):
            return f"{file.name}_{file.size}"
        if uploaded_files:
            current_files_map = {get_file_id(f): f for f in uploaded_files}
            current_ids = set(current_files_map.keys())
            st.session_state.merge_order = [fid for fid in st.session_state.merge_order if fid in current_ids]
            for fid in current_files_map:
                if fid not in st.session_state.merge_order:
                    st.session_state.merge_order.append(fid)
            st.write("### 🔢 Organize Files")
            def move_up(idx):
                if idx > 0:
                    st.session_state.merge_order[idx], st.session_state.merge_order[idx-1] = \
                    st.session_state.merge_order[idx-1], st.session_state.merge_order[idx]
            def move_down(idx):
                if idx < len(st.session_state.merge_order) - 1:
                    st.session_state.merge_order[idx], st.session_state.merge_order[idx+1] = \
                    st.session_state.merge_order[idx+1], st.session_state.merge_order[idx]
            for i, fid in enumerate(st.session_state.merge_order):
                f_obj = current_files_map[fid]
                c1, c2, c3 = st.columns([6, 1, 1])
                with c1: st.text(f"{i+1}. {f_obj.name}")
                with c2: st.button("⬆", key=f"up_{fid}", on_click=move_up, args=(i,), disabled=(i==0))
                with c3: st.button("⬇", key=f"down_{fid}", on_click=move_down, args=(i,), disabled=(i==len(st.session_state.merge_order)-1))
            if len(st.session_state.merge_order) >= 2:
                default_name = current_files_map[st.session_state.merge_order[0]].name.rsplit('.', 1)[0] + "_merged"
                output_name_input = st.text_input("Output Filename", value=default_name)
                final_output_name = f"{output_name_input}.pdf" if not output_name_input.lower().endswith('.pdf') else output_name_input
                if st.button("Merge PDFs Now"):
                    with st.spinner("Merging..."):
                        pdf_paths = []
                        try:
                            for fid in st.session_state.merge_order:
                                pdf_paths.append(save_uploaded(current_files_map[fid]))
                            merger = PDFMerger()
                            out_path = merger.merge(pdf_paths)
                            with open(out_path, "rb") as f:
                                st.download_button(label=f"⬇ Download {final_output_name}", data=f, file_name=final_output_name, mime="application/pdf")
                            if os.path.exists(out_path): os.remove(out_path)
                        except Exception as e: st.error(f"Merge Error: {e}")
                        finally:
                            for p in pdf_paths:
                                if os.path.exists(p): os.remove(p)
            else: st.info("Upload at least 2 files.")
        else:
            st.session_state.merge_order = []
            st.warning("Please upload at least 2 PDF files")

    elif tool_option == "Split PDF":
        st.subheader("Split PDF / Extract Pages")
        uploaded = st.file_uploader("Upload PDF to Split", type=["pdf"])
        if uploaded:
            try:
                import fitz
                with fitz.open(stream=uploaded.read(), filetype="pdf") as doc:
                    num_pages = len(doc)
                uploaded.seek(0)
                st.info(f"Total Pages: {num_pages}")
                page_range_str = st.text_input("Enter Page Numbers (e.g., 1, 3-5, 8)")
                if st.button("Extract Pages"):
                    if not page_range_str.strip(): st.error("Please enter a page range.")
                    else:
                        with st.spinner("Extracting..."):
                            tmp_path = save_uploaded(uploaded)
                            try:
                                selected_pages = []
                                parts = page_range_str.split(',')
                                for part in parts:
                                    part = part.strip()
                                    if '-' in part:
                                        start, end = part.split('-')
                                        selected_pages.extend(range(int(start)-1, int(end)))
                                    else: selected_pages.append(int(part)-1)
                                selected_pages = sorted(list(set(selected_pages)))
                                splitter = PDFSplitter()
                                out_path = splitter.extract_pages(tmp_path, selected_pages)
                                output_name = f"{os.path.splitext(uploaded.name)[0]}_extracted.pdf"
                                with open(out_path, "rb") as f:
                                    st.download_button("Download Extracted PDF", f, file_name=output_name, mime="application/pdf")
                                os.remove(out_path)
                            except Exception as e: st.error(f"Error: {e}")
                            finally:
                                if os.path.exists(tmp_path): os.remove(tmp_path)
            except Exception as e: st.error(f"Error: {e}")

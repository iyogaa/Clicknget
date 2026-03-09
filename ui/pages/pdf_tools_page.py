import streamlit as st
import streamlit.components.v1 as components
import os
import tempfile
import io
import fitz
from core.converters.pdf_processor import process_pdf
from core.converters.document_converter import WordToPDF, ExcelToPDF, ImageToPDF, PDFCompressor, PDFMerger, PDFSplitter


def save_uploaded(uploaded):
    suffix = os.path.splitext(uploaded.name)[1]
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(uploaded.getbuffer())
        return tmp.name


def _render_header(title_plain, title_em, subtitle):
    components.html(f"""
<!DOCTYPE html>
<html>
<head>
<link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@300;600&family=Inter:wght@300;400;500&display=swap" rel="stylesheet">
<style>
    *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{ background: transparent; overflow: hidden; }}
    .scene {{
        position: relative;
        display: flex;
        flex-direction: column;
        align-items: flex-start;
        justify-content: center;
        min-height: 130px;
        padding: 1.75rem 0.5rem 1.5rem 0.25rem;
        overflow: hidden;
    }}
    .deep-glow {{
        position: absolute;
        width: 480px; height: 200px;
        top: 50%; left: 30%;
        transform: translate(-50%, -50%);
        background: radial-gradient(ellipse at center,
            rgba(80, 100, 220, 0.06) 0%,
            rgba(50, 70, 180, 0.03) 45%,
            transparent 70%);
        pointer-events: none;
        filter: blur(20px);
        animation: glowBreathe 9s ease-in-out infinite alternate;
    }}
    #starCanvas {{
        position: absolute;
        inset: 0;
        width: 100%;
        height: 100%;
        pointer-events: none;
    }}
    .header-content {{ position: relative; z-index: 2; }}
    .eyebrow {{
        font-family: 'Inter', sans-serif;
        font-size: 0.62rem;
        font-weight: 500;
        letter-spacing: 0.2em;
        text-transform: uppercase;
        color: rgba(99, 118, 255, 0.85);
        margin-bottom: 0.55rem;
        opacity: 0;
        animation: fadeUp 0.7s cubic-bezier(0.16, 1, 0.3, 1) forwards 0.2s;
    }}
    .page-title {{
        font-family: 'Cormorant Garamond', Georgia, serif;
        font-size: clamp(1.7rem, 4vw, 2.2rem);
        font-weight: 300;
        letter-spacing: 0.01em;
        line-height: 1.15;
        color: #E4E8F5;
        opacity: 0;
        animation: fadeUp 0.9s cubic-bezier(0.16, 1, 0.3, 1) forwards 0.4s;
    }}
    .page-title em {{
        font-style: italic;
        font-weight: 600;
        background: linear-gradient(120deg, #9BAEFF 0%, #C8CFFF 45%, #7DD6EA 100%);
        background-size: 220% auto;
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        animation: shimmer 6s linear 1.5s infinite;
    }}
    .page-subtitle {{
        font-family: 'Inter', sans-serif;
        font-size: 0.8rem;
        font-weight: 300;
        color: rgba(130, 142, 172, 0.75);
        letter-spacing: 0.03em;
        margin-top: 0.5rem;
        opacity: 0;
        animation: fadeUp 0.9s cubic-bezier(0.16, 1, 0.3, 1) forwards 0.65s;
    }}
    .rule-bottom {{
        position: relative;
        z-index: 2;
        width: 0;
        height: 1px;
        background: linear-gradient(90deg, rgba(99,118,255,0.4), rgba(6,182,212,0.35), transparent);
        margin-top: 1.4rem;
        animation: expandRule 1.4s cubic-bezier(0.16, 1, 0.3, 1) forwards 0.9s;
    }}
    @keyframes fadeUp {{
        from {{ opacity: 0; transform: translateY(14px); }}
        to   {{ opacity: 1; transform: translateY(0); }}
    }}
    @keyframes expandRule {{
        from {{ width: 0; }}
        to   {{ width: 320px; }}
    }}
    @keyframes shimmer {{
        0%   {{ background-position: 0% center; }}
        100% {{ background-position: 220% center; }}
    }}
    @keyframes glowBreathe {{
        from {{ opacity: 0.7; transform: translate(-50%, -50%) scale(1); }}
        to   {{ opacity: 1;   transform: translate(-50%, -50%) scale(1.1); }}
    }}
</style>
</head>
<body>
<div class="scene">
    <div class="deep-glow"></div>
    <canvas id="starCanvas"></canvas>
    <div class="header-content">
        <p class="eyebrow">PDF Tools</p>
        <h1 class="page-title"><em>{title_em}</em> {title_plain}</h1>
        <p class="page-subtitle">{subtitle}</p>
        <div class="rule-bottom"></div>
    </div>
</div>
<script>
    const canvas = document.getElementById('starCanvas');
    const ctx = canvas.getContext('2d');
    canvas.width  = canvas.offsetWidth  || 800;
    canvas.height = canvas.offsetHeight || 130;
    const W = canvas.width, H = canvas.height;
    const stars = [];
    for (let i = 0; i < 90; i++) stars.push({{ x: Math.random()*W, y: Math.random()*H, r: Math.random()*0.55+0.15, baseAlpha: Math.random()*0.16+0.05, twinkleSpeed: Math.random()*0.004+0.001, twinkleOffset: Math.random()*Math.PI*2, tier: 1 }});
    for (let i = 0; i < 25; i++) stars.push({{ x: Math.random()*W, y: Math.random()*H, r: Math.random()*0.75+0.5,  baseAlpha: Math.random()*0.2+0.09,  twinkleSpeed: Math.random()*0.003+0.001, twinkleOffset: Math.random()*Math.PI*2, tier: 2 }});
    for (let i = 0; i < 5;  i++) stars.push({{ x: Math.random()*W, y: Math.random()*H, r: Math.random()*0.9+0.9,   baseAlpha: Math.random()*0.25+0.1,   twinkleSpeed: Math.random()*0.002+0.0005, twinkleOffset: Math.random()*Math.PI*2, tier: 3 }});
    let t = 0;
    function draw() {{
        ctx.clearRect(0, 0, W, H);
        t += 0.016;
        for (const s of stars) {{
            const alpha = s.baseAlpha * (0.6 + 0.4 * Math.sin(t * s.twinkleSpeed * 60 + s.twinkleOffset));
            if (s.tier === 3) {{
                const grd = ctx.createRadialGradient(s.x, s.y, 0, s.x, s.y, s.r*3.5);
                grd.addColorStop(0, `rgba(190,200,255,${{alpha}})`);
                grd.addColorStop(1, `rgba(190,200,255,0)`);
                ctx.beginPath(); ctx.arc(s.x, s.y, s.r*3.5, 0, Math.PI*2);
                ctx.fillStyle = grd; ctx.fill();
            }}
            ctx.beginPath(); ctx.arc(s.x, s.y, s.r, 0, Math.PI*2);
            ctx.fillStyle = `rgba(${{s.tier===1?'210,215,240':'200,210,255'}},${{alpha}})`;
            ctx.fill();
        }}
        requestAnimationFrame(draw);
    }}
    draw();
</script>
</body>
</html>
""", height=150)


def run_pdf_maker():

    _render_header("Flattening Tool", "PDF", "Convert interactive PDF forms to flat, static documents")

    tab1, tab2 = st.tabs(["Upload & Process", "Instructions"])

    with tab1:
        uploaded = st.file_uploader("Upload PDF", type=["pdf"], key="pdf_maker_upload", label_visibility="collapsed")

        if uploaded:
            col1, col2, col3 = st.columns([2, 1.5, 2])
            with col2:
                if st.button("Flatten PDF", use_container_width=True):
                    st.info("Processing PDF...")
                    temp_input = "uploaded.pdf"
                    with open(temp_input, "wb") as f:
                        f.write(uploaded.read())
                    output_file = "flattened.pdf"
                    process_pdf(temp_input, output_file)
                    with open(output_file, "rb") as f:
                        st.download_button(
                            "Download Flattened PDF", f,
                            file_name="flattened.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )
                    os.remove(temp_input)
                    os.remove(output_file)
                    st.success("PDF flattened successfully!")

    with tab2:
        st.markdown("""
        **What is PDF Flattening?**
        Removes interactive form fields, buttons, and annotations — converting all content to static,
        non-editable elements. Reduces file size and prevents accidental modifications.

        **When to Use**
        Filling out forms ready for distribution, archiving important documents, preparing PDFs for printing,
        or standardizing document format.

        **Steps:** Upload your interactive PDF → Click Flatten PDF → Download the flattened version.
        """)


def run_pdf_play():

    _render_header("Conversion Tools", "PDF", "Transform and manipulate PDF documents with precision")

    tool_option = st.selectbox(
        "Choose PDF Tool:",
        ["Word to PDF", "Excel to PDF", "Images to PDF", "Merge PDFs", "Split PDF", "Compress PDF"],
        label_visibility="collapsed"
    )

    st.markdown("")

    if tool_option == "Word to PDF":
        st.subheader("Word to PDF Converter")
        uploaded = st.file_uploader("Upload Word Document", type=["doc", "docx"], key="word_to_pdf", label_visibility="collapsed")
        if uploaded and st.button("Convert to PDF", use_container_width=True):
            with st.spinner("Converting..."):
                tmp_path = save_uploaded(uploaded)
                try:
                    converter = WordToPDF()
                    out_path = converter.convert(tmp_path)
                    with open(out_path, "rb") as f:
                        st.download_button("Download PDF", f, file_name=f"{uploaded.name}.pdf", mime="application/pdf", use_container_width=True)
                    st.success("Conversion successful!")
                    os.remove(out_path)
                except Exception as e:
                    st.error(f"Error: {e}")
                finally:
                    os.remove(tmp_path)

    elif tool_option == "Excel to PDF":
        st.subheader("Excel to PDF Converter")
        uploaded = st.file_uploader("Upload Excel File", type=["xls", "xlsx"], key="excel_to_pdf", label_visibility="collapsed")
        if uploaded and st.button("Convert to PDF", use_container_width=True):
            with st.spinner("Converting..."):
                tmp_path = save_uploaded(uploaded)
                try:
                    converter = ExcelToPDF()
                    out_path = converter.convert(tmp_path)
                    with open(out_path, "rb") as f:
                        st.download_button("Download PDF", f, file_name=f"{uploaded.name}.pdf", mime="application/pdf", use_container_width=True)
                    st.success("Conversion successful!")
                    os.remove(out_path)
                except Exception as e:
                    st.error(f"Error: {e}")
                finally:
                    os.remove(tmp_path)

    elif tool_option == "Images to PDF":
        st.subheader("Images to PDF Converter")
        uploaded_files = st.file_uploader("Upload Images", type=["jpg", "jpeg", "png"], accept_multiple_files=True, key="images_to_pdf", label_visibility="collapsed")
        if uploaded_files and st.button("Create PDF", use_container_width=True):
            with st.spinner("Converting..."):
                img_paths = [save_uploaded(up) for up in uploaded_files]
                try:
                    converter = ImageToPDF()
                    out_path = converter.convert(img_paths)
                    with open(out_path, "rb") as f:
                        st.download_button("Download PDF", f, file_name="images.pdf", mime="application/pdf", use_container_width=True)
                    st.success("PDF created successfully!")
                    if os.path.exists(out_path): os.remove(out_path)
                except Exception as e:
                    st.error(f"Error: {e}")
                finally:
                    for p in img_paths:
                        if os.path.exists(p): os.remove(p)

    elif tool_option == "Compress PDF":
        st.subheader("PDF Compressor")
        uploaded = st.file_uploader("Upload PDF", type=["pdf"], key="compress_pdf", label_visibility="collapsed")
        level = st.select_slider("Compression Level", options=["Low", "Medium", "High"], value="Medium")
        if uploaded and st.button("Compress", use_container_width=True):
            with st.spinner("Compressing..."):
                tmp_path = save_uploaded(uploaded)
                try:
                    compressor = PDFCompressor()
                    out_path = compressor.compress(tmp_path, level.lower())
                    original_size = os.path.getsize(tmp_path)
                    new_size = os.path.getsize(out_path)
                    reduction = (1 - new_size / original_size) * 100
                    st.success(f"Reduced by {reduction:.1f}% ({original_size/(1024*1024):.2f}MB to {new_size/(1024*1024):.2f}MB)")
                    with open(out_path, "rb") as f:
                        st.download_button("Download Compressed PDF", f, file_name=f"compressed_{uploaded.name}", mime="application/pdf", use_container_width=True)
                    os.remove(out_path)
                except Exception as e:
                    st.error(f"Error: {e}")
                finally:
                    os.remove(tmp_path)

    elif tool_option == "Merge PDFs":
        st.subheader("Merge Multiple PDFs")
        uploaded_files = st.file_uploader("Upload PDF files", type=["pdf"], accept_multiple_files=True, key="merge_pdf", label_visibility="collapsed")

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

            def move_up(idx):
                if idx > 0:
                    st.session_state.merge_order[idx], st.session_state.merge_order[idx-1] = \
                    st.session_state.merge_order[idx-1], st.session_state.merge_order[idx]

            def move_down(idx):
                if idx < len(st.session_state.merge_order) - 1:
                    st.session_state.merge_order[idx], st.session_state.merge_order[idx+1] = \
                    st.session_state.merge_order[idx+1], st.session_state.merge_order[idx]

            st.markdown("**Organize Files**")
            for i, fid in enumerate(st.session_state.merge_order):
                f_obj = current_files_map[fid]
                c1, c2, c3 = st.columns([6, 1, 1])
                with c1: st.text(f"{i+1}. {f_obj.name}")
                with c2: st.button("up", key=f"up_{fid}", on_click=move_up, args=(i,), disabled=(i == 0))
                with c3: st.button("dn", key=f"dn_{fid}", on_click=move_down, args=(i,), disabled=(i == len(st.session_state.merge_order)-1))

            if len(st.session_state.merge_order) >= 2:
                default_name = current_files_map[st.session_state.merge_order[0]].name.rsplit('.', 1)[0] + "_merged"
                output_name_input = st.text_input("Output Filename", value=default_name)
                final_output_name = f"{output_name_input}.pdf" if not output_name_input.lower().endswith('.pdf') else output_name_input
                if st.button("Merge PDFs", use_container_width=True):
                    with st.spinner("Merging..."):
                        pdf_paths = []
                        try:
                            for fid in st.session_state.merge_order:
                                pdf_paths.append(save_uploaded(current_files_map[fid]))
                            merger = PDFMerger()
                            out_path = merger.merge(pdf_paths)
                            with open(out_path, "rb") as f:
                                st.download_button(label=f"Download {final_output_name}", data=f, file_name=final_output_name, mime="application/pdf", use_container_width=True)
                            st.success("PDFs merged successfully!")
                            if os.path.exists(out_path): os.remove(out_path)
                        except Exception as e:
                            st.error(f"Merge Error: {e}")
                        finally:
                            for p in pdf_paths:
                                if os.path.exists(p): os.remove(p)
            else:
                st.info("Upload at least 2 files to merge.")
        else:
            st.session_state.merge_order = []
            st.warning("Please upload at least 2 PDF files.")

    elif tool_option == "Split PDF":
        st.subheader("Split PDF / Extract Pages")
        uploaded = st.file_uploader("Upload PDF to Split", type=["pdf"], key="split_pdf", label_visibility="collapsed")
        if uploaded:
            try:
                with fitz.open(stream=uploaded.read(), filetype="pdf") as doc:
                    num_pages = len(doc)
                uploaded.seek(0)
                st.info(f"Total Pages: {num_pages}")
                page_range_str = st.text_input("Enter Page Numbers (e.g. 1, 3-5, 8)")
                if st.button("Extract Pages", use_container_width=True):
                    if not page_range_str.strip():
                        st.error("Please enter a page range.")
                    else:
                        with st.spinner("Extracting..."):
                            tmp_path = save_uploaded(uploaded)
                            try:
                                selected_pages = []
                                for part in page_range_str.split(','):
                                    part = part.strip()
                                    if '-' in part:
                                        start, end = part.split('-')
                                        selected_pages.extend(range(int(start)-1, int(end)))
                                    else:
                                        selected_pages.append(int(part)-1)
                                selected_pages = sorted(list(set(selected_pages)))
                                splitter = PDFSplitter()
                                out_path = splitter.extract_pages(tmp_path, selected_pages)
                                output_name = f"{os.path.splitext(uploaded.name)[0]}_extracted.pdf"
                                with open(out_path, "rb") as f:
                                    st.download_button("Download Extracted PDF", f, file_name=output_name, mime="application/pdf", use_container_width=True)
                                st.success("Pages extracted successfully!")
                                os.remove(out_path)
                            except Exception as e:
                                st.error(f"Error: {e}")
                            finally:
                                if os.path.exists(tmp_path): os.remove(tmp_path)
            except Exception as e:
                st.error(f"Error reading PDF: {e}")
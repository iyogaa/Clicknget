import fitz
import tempfile

def merge_pdfs(uploaded_files):
    merged = fitz.open()

    for file in uploaded_files:
        merged.insert_pdf(
            fitz.open(stream=file.read(), filetype="pdf")
        )

    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    temp_file.close()
    
    merged.save(temp_file.name)
    merged.close()

    return temp_file.name

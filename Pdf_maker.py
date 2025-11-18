import fitz  # PyMuPDF
from pdf2image import convert_from_path
import img2pdf
import os
import tempfile

# Step 1: Flatten checkboxes using PyMuPDF
def fix_checkboxes(input_path, output_path):
    doc = fitz.open(input_path)

    for page in doc:
        widgets = page.widgets()
        if not widgets:
            continue

        for w in widgets:
            if w.field_type != fitz.PDF_WIDGET_TYPE_CHECKBOX:
                continue

            rect = w.rect
            checked = (w.field_value in ["Yes", "On", "/Yes", "/On", "1", True])

            shape = fitz.Shape(page)
            shape.draw_rect(rect)

            if checked:
                x1, y1, x2, y2 = rect
                shape.draw_line((x1 + 2, (y1 + y2) / 2), ((x1 + x2) / 2, y1 + 2))
                shape.draw_line(((x1 + x2) / 2, y1 + 2), (x2 - 2, y2 - 2))

            shape.finish(color=(0, 0, 0), fill=None)
            shape.commit()

        page.clean_contents()

    doc.save(output_path)
    doc.close()

# Step 2: Convert to non-editable PDF using pdf2image + img2pdf
def convert_to_non_editable(input_path, output_path):
    images = convert_from_path(
        input_path,
        dpi=300,
        fmt="png",
        poppler_path=None  # Use system-installed Poppler (works on Streamlit Cloud)
    )

    temp_files = []
    for img in images:
        temp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        img.convert("RGB").save(temp.name, "PNG")
        temp_files.append(temp.name)

    with open(output_path, "wb") as f:
        f.write(img2pdf.convert(temp_files))

    for tmp in temp_files:
        os.remove(tmp)

# Step 3: Full pipeline
def process_pdf(input_file, output_file):
    fixed_pdf = "fixed_temp.pdf"
    fix_checkboxes(input_file, fixed_pdf)
    convert_to_non_editable(fixed_pdf, output_file)
    os.remove(fixed_pdf)

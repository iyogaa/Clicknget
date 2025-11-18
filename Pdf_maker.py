import fitz  # PyMuPDF
import os

# Step 1: Flatten checkboxes
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

# Step 2: Convert to non-editable PDF using compressed JPEG images
def convert_to_non_editable(input_path, output_path):
    doc = fitz.open(input_path)
    new_pdf = fitz.open()

    for page in doc:
        pix = page.get_pixmap(dpi=150)  # Lower DPI for smaller size
        rect = fitz.Rect(0, 0, pix.width, pix.height)
        new_page = new_pdf.new_page(width=pix.width, height=pix.height)
        new_page.insert_image(rect, stream=pix.tobytes("jpeg"))  # No compress arg needed

    new_pdf.save(output_path, deflate=True)  # Enable PDF-level compression
    new_pdf.close()
    doc.close()

# Step 3: Full pipeline
def process_pdf(input_file, output_file):
    fixed_pdf = "fixed_temp.pdf"
    fix_checkboxes(input_file, fixed_pdf)
    convert_to_non_editable(fixed_pdf, output_file)
    os.remove(fixed_pdf)

import fitz
from pdf2image import convert_from_path
import img2pdf
import os

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


def convert_to_non_editable(input_path, output_path):
    # Streamlit Cloud already has Poppler installed
    images = convert_from_path(
        input_path,
        dpi=300,
        fmt="png"
    )

    temp_files = []
    for i, img in enumerate(images):
        temp_file = f"page_{i}.png"
        img = img.convert("RGB")
        img.save(temp_file, "PNG")
        temp_files.append(temp_file)

    with open(output_path, "wb") as f:
        f.write(img2pdf.convert(temp_files))

    for tmp in temp_files:
        os.remove(tmp)


def process_pdf(input_file, output_file):
    fixed_pdf = "fixed_temp.pdf"
    fix_checkboxes(input_file, fixed_pdf)
    convert_to_non_editable(fixed_pdf, output_file)
    os.remove(fixed_pdf)

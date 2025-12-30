import os
import io
import fitz  # pymupdf
import pandas as pd
from PIL import Image
import img2pdf
import shutil

# Try to import optional dependencies
try:
    import mammoth
except ImportError:
    mammoth = None

try:
    from xhtml2pdf import pisa
except ImportError:
    pisa = None

try:
    import pytesseract
except ImportError:
    pytesseract = None

try:
    from pdf2image import convert_from_path
except ImportError:
    convert_from_path = None

class WordToPDF:
    def __init__(self):
        if not mammoth:
            raise ImportError("mammoth library is required for Word to PDF conversion. Please install it.")
        if not pisa:
            raise ImportError("xhtml2pdf library is required for Word to PDF conversion. Please install it.")

    def convert(self, input_path):
        """Convert Word docx to PDF via HTML intermediate."""
        output_path = input_path.rsplit('.', 1)[0] + ".pdf"
        
        with open(input_path, "rb") as docx_file:
            result = mammoth.convert_to_html(docx_file)
            html = result.value
            
        # Add basic styling to ensure it looks decent
        html = f"""
        <html>
        <head>
            <style>
                body {{ font-family: 'Helvetica', sans-serif; padding: 20px; }}
                table {{ border-collapse: collapse; width: 100%; }}
                td, th {{ border: 1px solid #ddd; padding: 8px; }}
                img {{ max-width: 100%; height: auto; }}
            </style>
        </head>
        <body>
            {html}
        </body>
        </html>
        """
        
        with open(output_path, "wb") as pdf_file:
            pisa_status = pisa.CreatePDF(html, dest=pdf_file)
            
        if pisa_status.err:
            raise Exception("PDF generation failed")
            
        return output_path

class ExcelToPDF:
    def __init__(self):
        if not pisa:
            raise ImportError("xhtml2pdf library is required for Excel to PDF conversion. Please install it.")

    def convert(self, input_path):
        """Convert Excel sheets to PDF via HTML."""
        output_path = input_path.rsplit('.', 1)[0] + ".pdf"
        
        # Read all sheets
        xls = pd.ExcelFile(input_path)
        html_parts = []
        
        html_parts.append("""
        <html>
        <head>
            <style>
                body { font-family: 'Helvetica', sans-serif; padding: 20px; }
                h2 { color: #333; border-bottom: 2px solid #333; }
                table { border-collapse: collapse; width: 100%; margin-bottom: 20px; font-size: 10px; }
                td, th { border: 1px solid #ddd; padding: 4px; text-align: left; }
                th { background-color: #f2f2f2; }
            </style>
        </head>
        <body>
        """)
        
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            html_parts.append(f"<h2>Sheet: {sheet_name}</h2>")
            html_parts.append(df.to_html(index=False, na_rep=""))
            html_parts.append("<br><br>")
            
        html_parts.append("</body></html>")
        
        full_html = "".join(html_parts)
        
        # Close the Excel file handle explicitly
        xls.close()
        
        with open(output_path, "wb") as pdf_file:
            pisa_status = pisa.CreatePDF(full_html, dest=pdf_file)
            
        return output_path

class ImageToPDF:
    def __init__(self):
        pass

    def convert(self, image_files, output_filename="images.pdf"):
        """Convert list of image files to a single PDF."""
        # image_files can be a list of paths
        if not image_files:
            raise ValueError("No images provided")

        # Convert simple images
        # We use img2pdf for quality or PIL. img2pdf is great for JPEGs without re-encoding.
        # But for mixed types, let's use PIL to standardize.
        
        # Prepare valid paths
        img_paths = [img for img in image_files]
        
        # We can use img2pdf if they are all conformant, but it's picky.
        # Let's use fitz or PIL for maximum compatibility.
        # PIL approach:
        images = []
        for img_path in img_paths:
            img = Image.open(img_path)
            if img.mode != 'RGB':
                img = img.convert('RGB')
            images.append(img)

        output_path = os.path.join(os.path.dirname(img_paths[0]), output_filename)
        
        if len(images) == 1:
            images[0].save(output_path, "PDF", resolution=100.0)
        else:
            images[0].save(output_path, "PDF", save_all=True, append_images=images[1:], resolution=100.0)
            
        return output_path

class HTMLToPDF:
    def __init__(self):
        if not pisa:
             raise ImportError("xhtml2pdf library is required.")

    def convert(self, input_path):
        output_path = input_path.rsplit('.', 1)[0] + ".pdf"
        
        with open(input_path, "r", encoding="utf-8") as f:
            html_content = f.read()
            
        with open(output_path, "wb") as pdf_file:
            pisa.CreatePDF(html_content, dest=pdf_file)
            
        return output_path

class PDFOCR:
    def __init__(self):
        if not pytesseract:
            raise ImportError("pytesseract is required for OCR.")
        if not convert_from_path:
             raise ImportError("pdf2image is required for OCR.")

    def process(self, input_path):
        """Inject valid OCR into a PDF by converting pages to images and re-OCRing."""
        output_path = input_path.rsplit('.', 1)[0] + "_ocr.pdf"
        
        # This is a potentially heavy operation.
        # We will convert PDF to images, then use pytesseract to get PDF page from image, then merge.
        
        # Convert PDF to images
        # Poppler is needed for pdf2image. If not on system, this fails.
        # Assuming poppler is present or we fallback.
        try:
            images = convert_from_path(input_path)
        except Exception as e:
            raise Exception(f"Failed to convert PDF to images: {e}. Is Poppler installed?")

        ocr_pdf = fitz.open()
        
        for img in images:
            # Get PDF data from pytesseract
            pdf_bytes = pytesseract.image_to_pdf_or_hocr(img, extension='pdf')
            # Open as fitz doc
            img_doc = fitz.open("pdf", pdf_bytes)
            ocr_pdf.insert_pdf(img_doc)
        
        ocr_pdf.save(output_path)
        return output_path

class PDFCompressor:
    def __init__(self):
        pass

    def compress(self, input_path, level="medium"):
        """
        Compress PDF using PyMuPDF.
        Levels:
        - high: garbage=4, deflate=True, downsample images
        - medium: garbage=3, deflate=True
        - low: garbage=2, deflate=True
        """
        output_path = input_path.rsplit('.', 1)[0] + "_compressed.pdf"
        
        doc = fitz.open(input_path)
        
        # Compression settings
        if level == "high":
            # Re-compress images, clean contents
            doc.save(output_path, garbage=4, deflate=True, clean=True)
        elif level == "medium":
            doc.save(output_path, garbage=3, deflate=True)
        else: # low
            doc.save(output_path, garbage=2, deflate=True)
            
        return output_path

class PDFMerger:
    def __init__(self):
        pass

    def merge(self, input_paths):
        """Merge multiple PDFs into one."""
        if not input_paths:
            raise ValueError("No input files provided")
            
        merged = fitz.open()
        
        for path in input_paths:
            try:
                doc = fitz.open(path)
                merged.insert_pdf(doc)
            except Exception as e:
                raise Exception(f"Failed to merge {os.path.basename(path)}: {e}")
            
        # Use first file's directory and name structure for output
        output_path = input_paths[0].rsplit('.', 1)[0] + "_merged.pdf"
        
        merged.save(output_path)
        merged.close()
        
        return output_path

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

from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors



class WordToPDF:
    def __init__(self):
        if not mammoth:
            raise ImportError("mammoth library is required for Word conversion.")

    def convert(self, input_path):
        """Convert Word docx to PDF by extracting text."""
        output_path = input_path.rsplit('.', 1)[0] + ".pdf"
        
        with open(input_path, "rb") as docx_file:
            # We extract raw text because converting HTML->PDF properly without heavy libs is hard
            result = mammoth.extract_raw_text(docx_file)
            text = result.value
            
        doc = SimpleDocTemplate(output_path, pagesize=letter)
        styles = getSampleStyleSheet()
        story = []
        
        # Split by newlines and create paragraphs
        for line in text.split('\n'):
            if line.strip():
                story.append(Paragraph(line, styles["Normal"]))
                story.append(Spacer(1, 6))
                
        doc.build(story)
        return output_path

class ExcelToPDF:
    def __init__(self):
        pass

    def convert(self, input_path):
        """Convert Excel sheets to PDF tables using ReportLab."""
        output_path = input_path.rsplit('.', 1)[0] + ".pdf"
        
        # Read all sheets
        xls = pd.ExcelFile(input_path)
        
        doc = SimpleDocTemplate(output_path, pagesize=letter)
        elements = []
        styles = getSampleStyleSheet()
        
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            
            # Add Sheet Title
            elements.append(Paragraph(f"Sheet: {sheet_name}", styles["Heading2"]))
            elements.append(Spacer(1, 12))
            
            # Handle empty dataframes
            if df.empty:
                elements.append(Paragraph("(Empty Sheet)", styles["Normal"]))
                elements.append(Spacer(1, 24))
                continue
                
            # Convert DF to list of lists for Table
            # Add headers
            data = [df.columns.astype(str).tolist()] + df.astype(str).values.tolist()
            
            # Create Table
            t = Table(data)
            t.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ]))
            
            elements.append(t)
            elements.append(Spacer(1, 24))
        
        xls.close()
        doc.build(elements)
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
        pass

    def convert(self, input_path):
        """Convert HTML file to PDF (Basic text extraction)."""
        output_path = input_path.rsplit('.', 1)[0] + ".pdf"
        
        # Since we removed xhtml2pdf, we'll do basic text dump for now
        # to ensure cloud compatibility. 
        # For a real HTML parser without sys dependencies, it's complex.
        # We will strip tags and print text.
        
        with open(input_path, "r", encoding="utf-8") as f:
            html_content = f.read()
            
        # Very basic strip tags
        from io import StringIO
        from html.parser import HTMLParser

        class MLStripper(HTMLParser):
            def __init__(self):
                super().__init__()
                self.reset()
                self.strict = False
                self.convert_charrefs= True
                self.text = StringIO()
            def handle_data(self, d):
                self.text.write(d)
            def get_data(self):
                return self.text.getvalue()

        stripper = MLStripper()
        stripper.feed(html_content)
        text = stripper.get_data()
        
        doc = SimpleDocTemplate(output_path, pagesize=letter)
        styles = getSampleStyleSheet()
        story = []
        
        for line in text.split('\n'):
            if line.strip():
                story.append(Paragraph(line.strip(), styles["Normal"]))
                story.append(Spacer(1, 6))
                
        doc.build(story)
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

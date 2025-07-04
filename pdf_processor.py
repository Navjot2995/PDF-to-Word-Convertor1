import PyPDF2
from pdf2image import convert_from_path
import os
import tempfile
from PIL import Image
import io

class PDFProcessor:
    def __init__(self):
        self.temp_dir = None
    
    def has_extractable_text(self, pdf_path):
        """Check if PDF has extractable text content"""
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                
                # Check first few pages for text
                pages_to_check = min(3, len(pdf_reader.pages))
                total_text = ""
                
                for page_num in range(pages_to_check):
                    page = pdf_reader.pages[page_num]
                    text = page.extract_text()
                    total_text += text
                
                # Consider it has text if we extract meaningful content
                # (more than just whitespace and special characters)
                clean_text = ''.join(c for c in total_text if c.isalnum() or c.isspace())
                return len(clean_text.strip()) > 50
                
        except Exception as e:
            print(f"Error checking PDF text: {e}")
            return False
    
    def extract_text_with_formatting(self, pdf_path):
        """Extract text from PDF with basic formatting preservation"""
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                extracted_text = ""
                
                for page_num, page in enumerate(pdf_reader.pages):
                    if page_num > 0:
                        extracted_text += f"\n\n--- Page {page_num + 1} ---\n\n"
                    
                    page_text = page.extract_text()
                    
                    # Basic formatting preservation
                    lines = page_text.split('\n')
                    formatted_lines = []
                    
                    for line in lines:
                        line = line.strip()
                        if line:
                            formatted_lines.append(line)
                    
                    extracted_text += '\n'.join(formatted_lines)
                
                return extracted_text
                
        except Exception as e:
            raise Exception(f"Error extracting text from PDF: {e}")
    
    def convert_to_images(self, pdf_path):
        """Convert PDF pages to images for OCR processing"""
        try:
            # Convert PDF to images
            images = convert_from_path(
                pdf_path,
                dpi=300,  # High DPI for better OCR results
                fmt='RGB'
            )
            
            # Convert PIL Images to bytes for processing
            image_data = []
            for i, image in enumerate(images):
                # Convert to RGB if not already
                if image.mode != 'RGB':
                    image = image.convert('RGB')
                
                # Save to bytes
                img_byte_arr = io.BytesIO()
                image.save(img_byte_arr, format='PNG')
                img_byte_arr.seek(0)
                
                image_data.append({
                    'page_number': i + 1,
                    'image': image,
                    'bytes': img_byte_arr.getvalue()
                })
            
            return image_data
            
        except Exception as e:
            raise Exception(f"Error converting PDF to images: {e}")
    
    def get_pdf_info(self, pdf_path):
        """Get basic information about the PDF"""
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                
                info = {
                    'pages': len(pdf_reader.pages),
                    'title': pdf_reader.metadata.get('/Title', 'Unknown') if pdf_reader.metadata else 'Unknown',
                    'author': pdf_reader.metadata.get('/Author', 'Unknown') if pdf_reader.metadata else 'Unknown',
                    'creator': pdf_reader.metadata.get('/Creator', 'Unknown') if pdf_reader.metadata else 'Unknown'
                }
                
                return info
                
        except Exception as e:
            return {
                'pages': 0,
                'title': 'Unknown',
                'author': 'Unknown',
                'creator': 'Unknown',
                'error': str(e)
            }

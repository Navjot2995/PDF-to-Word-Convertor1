import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
import streamlit as st
import tempfile
from pathlib import Path
import traceback
from pdf_processor import PDFProcessor
from ocr_processor import OCRProcessor
from word_generator import WordGenerator

def main():
    st.title("📄 PDF to Word Converter")
    st.markdown("Convert your PDF documents to Word format with OCR support for handwritten text")
    
    # File upload section
    st.header("Upload PDF Document")
    uploaded_file = st.file_uploader(
        "Choose a PDF file",
        type=['pdf'],
        help="Upload a PDF document to convert to Word format. Supports both typed and handwritten documents."
    )
    
    if uploaded_file is not None:
        # Display file information
        st.success(f"File uploaded: {uploaded_file.name}")
        st.info(f"File size: {len(uploaded_file.getvalue()) / (1024*1024):.2f} MB")
        
        # Processing options
        st.header("Processing Options")
        col1, col2 = st.columns(2)
        
        with col1:
            force_ocr = st.checkbox(
                "Force OCR Processing",
                help="Check this if the PDF contains handwritten text or images with text"
            )
        
        with col2:
            preserve_formatting = st.checkbox(
                "Preserve Formatting",
                value=True,
                help="Attempt to preserve original document formatting"
            )
        
        # Convert button
        if st.button("Convert to Word", type="primary"):
            convert_pdf_to_word(uploaded_file, force_ocr, preserve_formatting)

def convert_pdf_to_word(uploaded_file, force_ocr, preserve_formatting):
    """Convert PDF to Word document"""
    
    # Create temporary directory for processing
    with tempfile.TemporaryDirectory() as temp_dir:
        try:
            # Save uploaded file
            pdf_path = os.path.join(temp_dir, "input.pdf")
            with open(pdf_path, "wb") as f:
                f.write(uploaded_file.getvalue())
            
            # Initialize processors
            pdf_processor = PDFProcessor()
            ocr_processor = OCRProcessor()
            word_generator = WordGenerator()
            
            # Progress tracking
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # Step 1: Analyze PDF
            status_text.text("Analyzing PDF structure...")
            progress_bar.progress(10)
            
            # Check if PDF has extractable text
            has_text = pdf_processor.has_extractable_text(pdf_path)
            
            if has_text and not force_ocr:
                st.info("📝 Document contains extractable text. Using direct text extraction.")
                
                # Step 2: Extract text directly
                status_text.text("Extracting text from PDF...")
                progress_bar.progress(30)
                
                extracted_content = pdf_processor.extract_text_with_formatting(pdf_path)
                progress_bar.progress(60)
                
            else:
                if force_ocr:
                    st.info("🔍 OCR processing requested. Converting PDF to images...")
                else:
                    st.info("🔍 No extractable text found. Using OCR processing...")
                
                # Step 2: Convert PDF to images
                status_text.text("Converting PDF pages to images...")
                progress_bar.progress(20)
                
                images = pdf_processor.convert_to_images(pdf_path)
                progress_bar.progress(40)
                
                # Step 3: OCR processing
                status_text.text("Performing OCR on images...")
                progress_bar.progress(50)
                
                extracted_content = ocr_processor.process_images(images)
                progress_bar.progress(80)
            
            # Step 4: Generate Word document
            status_text.text("Generating Word document...")
            progress_bar.progress(90)
            
            word_path = os.path.join(temp_dir, "output.docx")
            word_generator.create_document(extracted_content, word_path, preserve_formatting)
            
            progress_bar.progress(100)
            status_text.text("Conversion completed successfully!")
            
            # Step 5: Provide download
            st.success("✅ Conversion completed successfully!")
            
            # Display preview
            st.header("Preview")
            preview_text = extracted_content[:1000] + "..." if len(extracted_content) > 1000 else extracted_content
            st.text_area("Document Preview", preview_text, height=200)
            
            # Download button
            with open(word_path, "rb") as f:
                word_content = f.read()
            
            # Generate download filename
            original_name = Path(uploaded_file.name).stem
            download_name = f"{original_name}_converted.docx"
            
            st.download_button(
                label="📥 Download Word Document",
                data=word_content,
                file_name=download_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
            # Show statistics
            st.header("Conversion Statistics")
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric("Characters", len(extracted_content))
            
            with col2:
                words = len(extracted_content.split())
                st.metric("Words", words)
            
            with col3:
                lines = len(extracted_content.split('\n'))
                st.metric("Lines", lines)
                
        except Exception as e:
            st.error(f"❌ Error during conversion: {str(e)}")
            st.error("Please check your PDF file and try again.")
            
            # Show detailed error in expander for debugging
            with st.expander("Show detailed error"):
                st.code(traceback.format_exc())

if __name__ == "__main__":
    # Set page config
    st.set_page_config(
        page_title="PDF to Word Converter",
        page_icon="📄",
        layout="wide",
        initial_sidebar_state="collapsed"
    )
    
    # Add sidebar with instructions
    with st.sidebar:
        st.header("Instructions")
        st.markdown("""
        **How to use:**
        1. Upload a PDF file
        2. Choose processing options
        3. Click "Convert to Word"
        4. Download the converted document
        
        **Supported formats:**
        - PDF files with typed text
        - PDF files with handwritten text
        - Scanned PDF documents
        - Mixed content PDFs
        
        **Tips:**
        - For handwritten documents, enable "Force OCR Processing"
        - Large files may take longer to process
        - Check "Preserve Formatting" to maintain layout
        """)
        
        st.header("About")
        st.markdown("""
        This application uses:
        - **PyPDF2** for text extraction
        - **Tesseract OCR** for handwritten text
        - **python-docx** for Word generation
        - **Pillow** for image processing
        """)
    
    main()

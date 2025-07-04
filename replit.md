# PDF to Word Converter

## Overview

This is a Streamlit-based web application that converts PDF documents to Word format with OCR (Optical Character Recognition) support for handwritten text. The application can process both typed and handwritten documents, providing users with flexible conversion options and formatting preservation capabilities.

## System Architecture

The application follows a modular architecture with clear separation of concerns:

- **Frontend**: Streamlit web interface for user interaction
- **Processing Layer**: Utility modules for PDF processing, OCR, and Word generation
- **File Handling**: Temporary file management for processing uploaded documents

The system is designed as a single-page application with a straightforward user workflow: upload PDF → configure options → convert to Word.

## Key Components

### 1. Main Application (`app.py`)
- **Purpose**: Primary Streamlit interface and workflow orchestration
- **Key Features**:
  - File upload handling for PDF documents
  - Processing options configuration (Force OCR, Preserve Formatting)
  - User feedback and progress indication
  - Integration with processing utilities

### 2. PDF Processor (`utils/pdf_processor.py`)
- **Purpose**: Handles PDF document analysis and text extraction
- **Key Features**:
  - Text extractability detection (determines if OCR is needed)
  - Direct text extraction from PDFs with searchable text
  - Basic formatting preservation during extraction
  - Page-by-page processing with section markers

### 3. OCR Processor (`utils/ocr_processor.py`)
- **Purpose**: Optical Character Recognition for image-based and handwritten content
- **Key Features**:
  - Image preprocessing for better OCR accuracy
  - Tesseract OCR integration with optimized configuration
  - Image enhancement (contrast, sharpness, noise reduction)
  - Support for various image formats and quality levels

### 4. Word Generator (`utils/word_generator.py`)
- **Purpose**: Creates formatted Word documents from extracted text
- **Key Features**:
  - Document creation with proper formatting
  - Formatting preservation options
  - Page break handling for multi-page documents
  - Document metadata management

## Data Flow

1. **Upload**: User uploads PDF file through Streamlit interface
2. **Analysis**: PDF processor checks if document has extractable text
3. **Processing Decision**: 
   - If extractable text exists and Force OCR is not enabled: Direct text extraction
   - If no extractable text or Force OCR enabled: Convert to images → OCR processing
4. **Text Processing**: Extracted text is processed with formatting preservation
5. **Document Generation**: Word generator creates formatted DOCX file
6. **Download**: User receives converted Word document

## External Dependencies

### Core Libraries
- **Streamlit**: Web application framework
- **PyPDF2**: PDF text extraction and manipulation
- **pdf2image**: PDF to image conversion for OCR processing
- **pytesseract**: OCR engine wrapper
- **python-docx**: Word document generation
- **Pillow (PIL)**: Image processing and enhancement
- **OpenCV**: Advanced image preprocessing
- **NumPy**: Numerical operations for image processing

### System Dependencies
- **Tesseract OCR**: External OCR engine (requires system installation)
- **Poppler**: PDF rendering utilities (required by pdf2image)

## Deployment Strategy

The application is designed for deployment on Replit or similar cloud platforms:

- **Runtime**: Python 3.x environment
- **Dependencies**: Managed through requirements.txt or similar
- **System Requirements**: Tesseract OCR and Poppler utilities
- **File Handling**: Temporary file management for processing
- **Scalability**: Single-user sessions with temporary file cleanup

## Changelog

- July 04, 2025. Initial setup

## User Preferences

Preferred communication style: Simple, everyday language.
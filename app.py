import streamlit as st
from pathlib import Path
from pptx import Presentation
import fitz  # PyMuPDF
import pandas as pd
from openai import OpenAI

# Initialize OpenAI client
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

def extract_ppt_content(file_path):
    """Extract text content from PowerPoint files."""
    try:
        prs = Presentation(file_path)
        slide_contents = []
        notes = []

        for slide in prs.slides:
            # Extract text from shapes
            slide_text = []
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    slide_text.append(shape.text)
            
            slide_contents.append(' '.join(slide_text))
            
            # Extract notes
            if slide.has_notes_slide:
                notes_text = slide.notes_slide.notes_text_frame.text
                notes.append(notes_text)

        return {
            'filename': str(file_path),
            'slide_contents': slide_contents,
            'notes': notes,
            'file_type': 'ppt'
        }
    except Exception as e:
        st.error(f"Error processing {file_path}: {str(e)}")
        return None

def extract_pdf_content(file_path):
    """Extract text content from PDF files."""
    try:
        pdf_content = []
        
        # Open the PDF with PyMuPDF
        doc = fitz.open(file_path)
        
        for page in doc:
            # Extract text
            pdf_content.append(page.get_text())

        return {
            'filename': str(file_path),
            'slide_contents': pdf_content,
            'notes': [],  # PDFs don't have native notes
            'file_type': 'pdf'
        }
    except Exception as e:
        st.error(f"Error processing {file_path}: {str(e)}")
        return None

def main():
    st.set_page_config(
        page_title="Presentation Content Search",
        page_icon="🔍",
        layout="wide"
    )
    
    st.title("🔍 Presentation Content Search")
    st.write("Search through your presentations (PPT, PPTX, PDF)")

    # File uploader for testing
    uploaded_file = st.file_uploader("Upload a presentation", type=['ppt', 'pptx', 'pdf'])
    
    if uploaded_file:
        # Save the uploaded file temporarily
        temp_path = Path(f"temp_{uploaded_file.name}")
        with open(temp_path, 'wb') as f:
            f.write(uploaded_file.getbuffer())
        
        try:
            # Process the file based on its type
            if uploaded_file.name.lower().endswith(('.ppt', '.pptx')):
                content = extract_ppt_content(temp_path)
            else:
                content = extract_pdf_content(temp_path)
            
            if content:
                st.subheader("Extracted Content")
                st.json(content)
                
        finally:
            # Clean up temporary file
            temp_path.unlink()

if __name__ == "__main__":
    main()

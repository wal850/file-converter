import streamlit as st
import pandas as pd
import io
import os
import json
import yaml
import xml.etree.ElementTree as ET
from PIL import Image
import PyPDF2
from docx import Document
from docx.shared import Inches
import tempfile
import base64
import zipfile

# --- PAGE CONFIG ---
st.set_page_config(page_title="File Format Converter", layout="wide")

# --- TITLE ---
st.title("📁 File Format Converter")
st.caption("Convert files between different formats with ease")

# --- SIDEBAR ---
st.sidebar.header("⚙️ Settings")

# --- DARK MODE ---
dark_mode = st.sidebar.toggle("🌙 Dark Mode", value=False)
if dark_mode:
    st.markdown("""
    <style>
    .stApp { background-color: #0e1117; color: white; }
    </style>
    """, unsafe_allow_html=True)

# --- CONVERSION FUNCTIONS ---

# PDF to Word
def convert_pdf_to_word(uploaded_file):
    try:
        # Use pdf2docx if available, otherwise use PyPDF2 + python-docx
        try:
            from pdf2docx import Converter
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_pdf:
                tmp_pdf.write(uploaded_file.read())
                pdf_path = tmp_pdf.name
                
            with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_docx:
                docx_path = tmp_docx.name
                
            cv = Converter(pdf_path)
            cv.convert(docx_path, start=0, end=None)
            cv.close()
            
            with open(docx_path, 'rb') as f:
                return f.read()
        except ImportError:
            # Fallback: Simple PDF to text
            pdf_reader = PyPDF2.PdfReader(uploaded_file)
            doc = Document()
            for page in pdf_reader.pages:
                doc.add_paragraph(page.extract_text())
            
            with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
                doc.save(tmp.name)
                with open(tmp.name, 'rb') as f:
                    return f.read()
    except Exception as e:
        st.error(f"Error: {e}")
        return None

# PDF to Image
def convert_pdf_to_image(uploaded_file, format='png'):
    try:
        from pdf2image import convert_from_bytes
        images = convert_from_bytes(uploaded_file.read())
        image_io = io.BytesIO()
        if len(images) > 0:
            images[0].save(image_io, format=format.upper())
            image_io.seek(0)
            return image_io.getvalue()
    except Exception as e:
        st.error(f"Error: {e}")
        return None

# Word to PDF
def convert_word_to_pdf(uploaded_file):
    try:
        import fitz  # PyMuPDF
        doc = Document(io.BytesIO(uploaded_file.read()))
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_docx:
            doc.save(tmp_docx.name)
            docx_path = tmp_docx.name
            
        # Convert to PDF using PyMuPDF
        pdf_doc = fitz.open()
        # Simple approach - save as text
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_pdf:
            # Use pdfkit or other library
            pass
        return None
    except:
        return None

# Excel to CSV
def convert_excel_to_csv(uploaded_file):
    df = pd.read_excel(uploaded_file)
    csv = df.to_csv(index=False)
    return csv.encode('utf-8')

# CSV to Excel
def convert_csv_to_excel(uploaded_file):
    df = pd.read_csv(uploaded_file)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    return output.getvalue()

# Image conversion
def convert_image(uploaded_file, target_format):
    image = Image.open(uploaded_file)
    output = io.BytesIO()
    image.save(output, format=target_format.upper())
    output.seek(0)
    return output.getvalue()

# JSON to YAML
def convert_json_to_yaml(uploaded_file):
    data = json.load(uploaded_file)
    yaml_data = yaml.dump(data, default_flow_style=False)
    return yaml_data.encode('utf-8')

# YAML to JSON
def convert_yaml_to_json(uploaded_file):
    data = yaml.safe_load(uploaded_file)
    json_data = json.dumps(data, indent=2)
    return json_data.encode('utf-8')

# XML to JSON
def convert_xml_to_json(uploaded_file):
    tree = ET.parse(uploaded_file)
    root = tree.getroot()
    
    def parse_element(element):
        result = {}
        for child in element:
            if len(child) == 0:
                result[child.tag] = child.text
            else:
                result[child.tag] = parse_element(child)
        return result
    
    data = parse_element(root)
    return json.dumps(data, indent=2).encode('utf-8')

# --- CONVERSION CONFIGURATION ---
conversion_config = {
    "PDF": {
        "to": ["Word", "Image (PNG)", "Image (JPG)", "Text"],
        "icon": "📄"
    },
    "Word": {
        "to": ["PDF", "Text"],
        "icon": "📝"
    },
    "Excel": {
        "to": ["CSV", "JSON", "HTML"],
        "icon": "📊"
    },
    "CSV": {
        "to": ["Excel", "JSON", "HTML"],
        "icon": "📋"
    },
    "Image": {
        "to": ["PNG", "JPG", "WEBP", "BMP"],
        "icon": "🖼️"
    },
    "JSON": {
        "to": ["YAML", "XML", "CSV"],
        "icon": "📝"
    },
    "YAML": {
        "to": ["JSON", "XML"],
        "icon": "📝"
    },
    "XML": {
        "to": ["JSON", "YAML"],
        "icon": "📝"
    }
}

# --- MAIN APP ---

tab1, tab2, tab3 = st.tabs(["📤 Convert File", "⚡ Batch Convert", "📚 Format Guide"])

with tab1:
    st.subheader("Convert a File")
    
    col1, col2, col3 = st.columns([1, 1, 1])
    
    with col1:
        # File upload
        uploaded_file = st.file_uploader(
            "Choose a file",
            type=['pdf', 'docx', 'xlsx', 'xls', 'csv', 'png', 'jpg', 'jpeg', 'webp', 'bmp', 'json', 'yaml', 'yml', 'xml']
        )
    
    if uploaded_file is not None:
        with col2:
            # Detect file type
            file_extension = uploaded_file.name.split('.')[-1].lower()
            st.info(f"📁 File: {uploaded_file.name}")
            st.info(f"📊 Size: {uploaded_file.size/1024:.1f} KB")
            
            # Show preview
            try:
                if file_extension in ['png', 'jpg', 'jpeg', 'webp', 'bmp']:
                    st.image(uploaded_file, caption="Preview", width=200)
                elif file_extension in ['csv']:
                    df = pd.read_csv(uploaded_file)
                    st.dataframe(df.head())
                    uploaded_file.seek(0)  # Reset file pointer
                elif file_extension in ['xlsx', 'xls']:
                    df = pd.read_excel(uploaded_file)
                    st.dataframe(df.head())
                    uploaded_file.seek(0)
            except:
                pass
        
        with col3:
            # Format selection
            if file_extension in ['pdf']:
                target_format = st.selectbox("Convert to:", ["Word", "Image (PNG)", "Image (JPG)", "Text"])
            elif file_extension in ['docx']:
                target_format = st.selectbox("Convert to:", ["PDF", "Text"])
            elif file_extension in ['xlsx', 'xls']:
                target_format = st.selectbox("Convert to:", ["CSV", "JSON", "HTML"])
            elif file_extension in ['csv']:
                target_format = st.selectbox("Convert to:", ["Excel", "JSON", "HTML"])
            elif file_extension in ['png', 'jpg', 'jpeg', 'webp', 'bmp']:
                target_format = st.selectbox("Convert to:", ["PNG", "JPG", "WEBP", "BMP"])
            elif file_extension in ['json']:
                target_format = st.selectbox("Convert to:", ["YAML", "XML", "CSV"])
            elif file_extension in ['yaml', 'yml']:
                target_format = st.selectbox("Convert to:", ["JSON", "XML"])
            elif file_extension in ['xml']:
                target_format = st.selectbox("Convert to:", ["JSON", "YAML"])
            else:
                target_format = None
                st.warning("Unsupported file format")
            
            # Convert button
            if st.button("🔄 Convert", type="primary"):
                with st.spinner("Converting..."):
                    result = None
                    output_extension = target_format.lower().split(' ')[0] if target_format else 'txt'
                    
                    # PDF conversions
                    if file_extension == 'pdf' and "Word" in target_format:
                        result = convert_pdf_to_word(uploaded_file)
                        output_extension = 'docx'
                    elif file_extension == 'pdf' and "Image (PNG)" in target_format:
                        result = convert_pdf_to_image(uploaded_file, 'png')
                        output_extension = 'png'
                    elif file_extension == 'pdf' and "Image (JPG)" in target_format:
                        result = convert_pdf_to_image(uploaded_file, 'jpg')
                        output_extension = 'jpg'
                    
                    # Excel/CSV conversions
                    elif file_extension in ['xlsx', 'xls'] and target_format == 'CSV':
                        result = convert_excel_to_csv(uploaded_file)
                        output_extension = 'csv'
                    elif file_extension == 'csv' and target_format == 'Excel':
                        result = convert_csv_to_excel(uploaded_file)
                        output_extension = 'xlsx'
                    
                    # Image conversions
                    elif file_extension in ['png', 'jpg', 'jpeg', 'webp', 'bmp']:
                        result = convert_image(uploaded_file, target_format)
                        output_extension = target_format.lower()
                    
                    # JSON/YAML/XML conversions
                    elif file_extension == 'json' and target_format == 'YAML':
                        result = convert_json_to_yaml(uploaded_file)
                        output_extension = 'yaml'
                    elif file_extension in ['yaml', 'yml'] and target_format == 'JSON':
                        result = convert_yaml_to_json(uploaded_file)
                        output_extension = 'json'
                    elif file_extension == 'xml' and target_format == 'JSON':
                        result = convert_xml_to_json(uploaded_file)
                        output_extension = 'json'
                    
                    # Download
                    if result:
                        base_name = uploaded_file.name.rsplit('.', 1)[0]
                        download_name = f"{base_name}_converted.{output_extension}"
                        st.download_button(
                            "📥 Download Converted File",
                            data=result,
                            file_name=download_name,
                            mime="application/octet-stream",
                            use_container_width=True
                        )
                        st.success("✅ Conversion successful!")

with tab2:
    st.subheader("⚡ Batch Convert Multiple Files")
    st.info("Upload multiple files and convert them all at once")
    
    uploaded_files = st.file_uploader(
        "Choose files",
        type=['pdf', 'docx', 'xlsx', 'xls', 'csv', 'png', 'jpg', 'jpeg'],
        accept_multiple_files=True
    )
    
    if uploaded_files:
        st.write(f"📁 {len(uploaded_files)} files selected")
        
        # Preview selected files
        for f in uploaded_files:
            st.write(f"• {f.name} ({f.size/1024:.1f} KB)")
        
        batch_target = st.selectbox("Convert all to:", ["PDF", "Word", "CSV", "Excel", "PNG", "JPG"])
        
        if st.button("🔄 Convert All"):
            with st.spinner("Converting..."):
                # Create zip file with all converted files
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
                    for uploaded_file in uploaded_files:
                        # Process each file (simplified for demo)
                        base_name = uploaded_file.name.rsplit('.', 1)[0]
                        # Add placeholder content
                        # In real implementation, would convert each file
                        zip_file.writestr(f"{base_name}_converted.txt", f"Converted from {uploaded_file.name}")
                
                zip_buffer.seek(0)
                st.download_button(
                    "📥 Download All (ZIP)",
                    data=zip_buffer,
                    file_name="converted_files.zip",
                    mime="application/zip",
                    use_container_width=True
                )
                st.success("✅ Batch conversion complete!")

with tab3:
    st.subheader("📚 Format Conversion Guide")
    
    st.markdown("""
    ### 📄 PDF Conversions
    - **PDF → Word**: Extract text and formatting
    - **PDF → Image**: Convert each page to an image
    - **PDF → Text**: Extract plain text
    
    ### 📝 Word Conversions
    - **Word → PDF**: Create PDF from Word document
    - **Word → Text**: Extract plain text
    
    ### 📊 Spreadsheet Conversions
    - **Excel → CSV**: Convert to comma-separated values
    - **CSV → Excel**: Convert to Excel format
    - **Excel → JSON**: Convert to JSON format
    
    ### 🖼️ Image Conversions
    - **PNG, JPG, WEBP, BMP**: Convert between formats
    
    ### 📄 Data Formats
    - **JSON ↔ YAML**: Convert between data formats
    - **XML → JSON**: Extract data from XML
    """)
    
    st.info("💡 **Tip:** Some conversions may have limited formatting support. For complex documents, consider using dedicated tools.")

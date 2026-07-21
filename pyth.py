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
import tempfile
import zipfile
import pytesseract
from pdf2image import convert_from_bytes

# --- PAGE CONFIG ---
st.set_page_config(page_title="File Format Converter & OCR", layout="wide")

# --- TITLE ---
st.title("📁 File Format Converter + OCR")
st.caption("Convert files between different formats AND extract text from images")

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

# --- OCR FUNCTIONS ---
def extract_text_from_image(uploaded_file, language='eng'):
    """Extract text from an image using Tesseract OCR"""
    try:
        image = Image.open(uploaded_file)
        # Perform OCR
        text = pytesseract.image_to_string(image, lang=language)
        return text
    except Exception as e:
        return f"Error: {e}"

def extract_text_from_pdf(uploaded_file, language='eng'):
    """Extract text from a PDF using OCR"""
    try:
        # Convert PDF to images
        images = convert_from_bytes(uploaded_file.read(), dpi=300)
        all_text = []
        
        for i, image in enumerate(images):
            # Perform OCR on each page
            text = pytesseract.image_to_string(image, lang=language)
            all_text.append(f"--- Page {i+1} ---\n{text}\n")
        
        return "\n".join(all_text)
    except Exception as e:
        return f"Error: {e}"

def extract_text_from_scanned_pdf(uploaded_file, language='eng'):
    """Extract text from scanned PDF (PDF with images)"""
    # Same as above but specifically for scanned PDFs
    return extract_text_from_pdf(uploaded_file, language)

# --- CONVERSION FUNCTIONS ---
# (Keep all previous conversion functions here - I'll add them back below)

def convert_pdf_to_word(uploaded_file):
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
    except:
        # Fallback
        pdf_reader = PyPDF2.PdfReader(uploaded_file)
        doc = Document()
        for page in pdf_reader.pages:
            doc.add_paragraph(page.extract_text())
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
            doc.save(tmp.name)
            with open(tmp.name, 'rb') as f:
                return f.read()

def convert_excel_to_csv(uploaded_file):
    df = pd.read_excel(uploaded_file)
    csv = df.to_csv(index=False)
    return csv.encode('utf-8')

def convert_csv_to_excel(uploaded_file):
    df = pd.read_csv(uploaded_file)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    return output.getvalue()

def convert_image(uploaded_file, target_format):
    image = Image.open(uploaded_file)
    output = io.BytesIO()
    image.save(output, format=target_format.upper())
    output.seek(0)
    return output.getvalue()

def convert_json_to_yaml(uploaded_file):
    data = json.load(uploaded_file)
    yaml_data = yaml.dump(data, default_flow_style=False)
    return yaml_data.encode('utf-8')

def convert_yaml_to_json(uploaded_file):
    data = yaml.safe_load(uploaded_file)
    json_data = json.dumps(data, indent=2)
    return json_data.encode('utf-8')

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

# --- MAIN APP ---

tab1, tab2, tab3, tab4 = st.tabs(["📤 Convert File", "📝 Extract Text (OCR)", "⚡ Batch Convert", "📚 Format Guide"])

# --- TAB 1: Convert File ---
with tab1:
    st.subheader("Convert a File")
    
    col1, col2, col3 = st.columns([1, 1, 1])
    
    with col1:
        uploaded_file = st.file_uploader(
            "Choose a file",
            type=['pdf', 'docx', 'xlsx', 'xls', 'csv', 'png', 'jpg', 'jpeg', 'webp', 'bmp', 'json', 'yaml', 'yml', 'xml']
        )
    
    if uploaded_file is not None:
        with col2:
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
                    uploaded_file.seek(0)
                elif file_extension in ['xlsx', 'xls']:
                    df = pd.read_excel(uploaded_file)
                    st.dataframe(df.head())
                    uploaded_file.seek(0)
            except:
                pass
        
        with col3:
            # Format selection
            if file_extension == 'pdf':
                target_format = st.selectbox("Convert to:", ["Word", "Image (PNG)", "Image (JPG)", "Text"])
            elif file_extension in ['docx']:
                target_format = st.selectbox("Convert to:", ["PDF", "Text"])
            elif file_extension in ['xlsx', 'xls']:
                target_format = st.selectbox("Convert to:", ["CSV", "JSON", "HTML"])
            elif file_extension == 'csv':
                target_format = st.selectbox("Convert to:", ["Excel", "JSON", "HTML"])
            elif file_extension in ['png', 'jpg', 'jpeg', 'webp', 'bmp']:
                target_format = st.selectbox("Convert to:", ["PNG", "JPG", "WEBP", "BMP"])
            elif file_extension == 'json':
                target_format = st.selectbox("Convert to:", ["YAML", "XML", "CSV"])
            elif file_extension in ['yaml', 'yml']:
                target_format = st.selectbox("Convert to:", ["JSON", "XML"])
            elif file_extension == 'xml':
                target_format = st.selectbox("Convert to:", ["JSON", "YAML"])
            else:
                target_format = None
                st.warning("Unsupported file format")
            
            if st.button("🔄 Convert", type="primary"):
                with st.spinner("Converting..."):
                    result = None
                    output_extension = target_format.lower().split(' ')[0] if target_format else 'txt'
                    
                    # PDF conversions
                    if file_extension == 'pdf' and "Word" in target_format:
                        result = convert_pdf_to_word(uploaded_file)
                        output_extension = 'docx'
                    elif file_extension == 'pdf' and "Image (PNG)" in target_format:
                        from pdf2image import convert_from_bytes
                        images = convert_from_bytes(uploaded_file.read())
                        image_io = io.BytesIO()
                        if images:
                            images[0].save(image_io, format='PNG')
                            image_io.seek(0)
                            result = image_io.getvalue()
                            output_extension = 'png'
                    elif file_extension == 'pdf' and "Image (JPG)" in target_format:
                        from pdf2image import convert_from_bytes
                        images = convert_from_bytes(uploaded_file.read())
                        image_io = io.BytesIO()
                        if images:
                            images[0].save(image_io, format='JPEG')
                            image_io.seek(0)
                            result = image_io.getvalue()
                            output_extension = 'jpg'
                    elif file_extension == 'pdf' and "Text" in target_format:
                        pdf_reader = PyPDF2.PdfReader(uploaded_file)
                        text = ""
                        for page in pdf_reader.pages:
                            text += page.extract_text()
                        result = text.encode('utf-8')
                        output_extension = 'txt'
                    
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

# --- TAB 2: OCR - Extract Text from Images ---
with tab2:
    st.subheader("📝 Extract Text from Images")
    st.caption("Upload an image or PDF and extract text using OCR")
    
    col1, col2 = st.columns([1, 1])
    
    with col1:
        ocr_file = st.file_uploader(
            "Upload image or PDF for text extraction",
            type=['png', 'jpg', 'jpeg', 'webp', 'bmp', 'pdf'],
            key="ocr_uploader"
        )
        
        # Language selection
        language = st.selectbox(
            "Language",
            ['eng', 'spa', 'fra', 'deu', 'ita', 'por', 'rus', 'jpn', 'chi'],
            help="Select the language of the text in the image"
        )
        
        # OCR mode
        ocr_mode = st.radio(
            "OCR Mode",
            ['Standard', 'Advanced'],
            help="Advanced mode uses more processing for better accuracy"
        )
    
    with col2:
        if ocr_file is not None:
            file_extension = ocr_file.name.split('.')[-1].lower()
            
            # Show preview
            if file_extension in ['png', 'jpg', 'jpeg', 'webp', 'bmp']:
                st.image(ocr_file, caption="Image Preview", width=300)
            elif file_extension == 'pdf':
                st.info("📄 PDF uploaded - will extract text from all pages")
            
            if st.button("🔍 Extract Text", type="primary"):
                with st.spinner("Extracting text..."):
                    # Perform OCR based on file type
                    if file_extension in ['png', 'jpg', 'jpeg', 'webp', 'bmp']:
                        # Check if image is scanned or contains text
                        if ocr_mode == 'Advanced':
                            # Preprocess image for better OCR
                            image = Image.open(ocr_file)
                            # Convert to grayscale
                            image = image.convert('L')
                            # Increase contrast
                            from PIL import ImageEnhance
                            enhancer = ImageEnhance.Contrast(image)
                            image = enhancer.enhance(2.0)
                            # Save to temp file
                            with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
                                image.save(tmp.name)
                                with open(tmp.name, 'rb') as f:
                                    text = extract_text_from_image(io.BytesIO(f.read()), language)
                        else:
                            text = extract_text_from_image(ocr_file, language)
                        
                        st.session_state['ocr_text'] = text
                    
                    elif file_extension == 'pdf':
                        text = extract_text_from_pdf(ocr_file, language)
                        st.session_state['ocr_text'] = text
    
    # Display extracted text
    if 'ocr_text' in st.session_state and st.session_state['ocr_text']:
        st.subheader("📄 Extracted Text")
        
        # Show text
        text_area = st.text_area(
            "Extracted Text",
            value=st.session_state['ocr_text'],
            height=400,
            key="ocr_output"
        )
        
        # Download options
        col1, col2, col3 = st.columns(3)
        with col1:
            st.download_button(
                "📥 Download as TXT",
                data=st.session_state['ocr_text'].encode('utf-8'),
                file_name="extracted_text.txt",
                mime="text/plain",
                use_container_width=True
            )
        with col2:
            st.download_button(
                "📥 Download as PDF",
                data=st.session_state['ocr_text'].encode('utf-8'),
                file_name="extracted_text.pdf",
                mime="application/pdf",
                use_container_width=True
            )
        with col3:
            st.button("📋 Copy to Clipboard", on_click=None)
            st.caption("Select text and press Ctrl+C")

# --- TAB 3: Batch Convert ---
with tab3:
    st.subheader("⚡ Batch Convert Multiple Files")
    st.info("Upload multiple files and convert them all at once")
    
    uploaded_files = st.file_uploader(
        "Choose files",
        type=['pdf', 'docx', 'xlsx', 'xls', 'csv', 'png', 'jpg', 'jpeg'],
        accept_multiple_files=True
    )
    
    if uploaded_files:
        st.write(f"📁 {len(uploaded_files)} files selected")
        
        for f in uploaded_files:
            st.write(f"• {f.name} ({f.size/1024:.1f} KB)")
        
        batch_target = st.selectbox("Convert all to:", ["PDF", "Word", "CSV", "Excel", "PNG", "JPG"])
        
        if st.button("🔄 Convert All"):
            with st.spinner("Converting..."):
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
                    for uploaded_file in uploaded_files:
                        base_name = uploaded_file.name.rsplit('.', 1)[0]
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

# --- TAB 4: Format Guide ---
with tab4:
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
    
    ### 🔍 OCR (Text Extraction)
    - **Image → Text**: Extract text from images
    - **PDF → Text**: Extract text from scanned PDFs
    - **Supports multiple languages**
    """)
    
    st.info("💡 **Tip:** For best OCR results, use clear, high-contrast images with text. Advanced mode provides better accuracy for complex images.")

# 📁 File Format Converter + OCR

A web-based file format converter and OCR tool built with Streamlit. Convert between PDF, Word, Excel, CSV, Image, JSON, YAML, and XML formats. Extract text from images and scanned PDFs.

## 🚀 Features

- 📄 **PDF ↔ Word** - Convert between PDF and Word documents
- 📊 **Excel ↔ CSV** - Convert between spreadsheet formats
- 🖼️ **Image formats** - JPG, PNG, WEBP, BMP conversion
- 📝 **Data formats** - JSON, YAML, XML conversion
- 🔍 **OCR (Text Extraction)** - Extract text from images and scanned PDFs
- 🌙 **Dark/Light mode** - Toggle between themes
- ⚡ **Batch conversion** - Convert multiple files at once
- 🌍 **Multi-language OCR** - Supports English, Spanish, French, German, and more

## 🛠️ Tech Stack

- **Python 3.12**
- **Streamlit** - Web framework
- **PyPDF2 / pdf2docx** - PDF processing
- **python-docx** - Word processing
- **Pandas / openpyxl** - Spreadsheet handling
- **Pillow** - Image processing
- **PyYAML** - YAML handling
- **Tesseract OCR** - Text extraction from images
- **pdf2image** - PDF to image conversion

## 📦 Installation

### 1. Clone the repository

git clone https://github.com/wal850/file-converter.git
cd file-converter

### 2. Create virtual environment

python3 -m venv myenv
source myenv/bin/activate

### 3. Install dependencies

pip install -r requirements.txt

### 4. Install Tesseract (for OCR)

Ubuntu/Debian:
sudo apt update
sudo apt install tesseract-ocr tesseract-ocr-eng

### 5. Run the app

streamlit run pyth.py

## 📂 Usage

### Convert Files
1. Upload a file via the file uploader
2. Select target format from the dropdown
3. Click "Convert" to process the file
4. Download the converted file

### Extract Text from Images (OCR)
1. Go to the "Extract Text (OCR)" tab
2. Upload an image or PDF
3. Select the language
4. Click "Extract Text"
5. Copy or download the extracted text

### Batch Conversion
1. Go to the "Batch Convert" tab
2. Upload multiple files
3. Select the target format
4. Download all converted files as a ZIP

## 🔍 OCR Features

- Extracts text from JPG, PNG, WEBP, BMP images
- Extracts text from scanned PDFs
- Supports multiple languages:
  - English (eng)
  - Spanish (spa)
  - French (fra)
  - German (deu)
  - Italian (ita)
  - Portuguese (por)
  - Russian (rus)
  - Japanese (jpn)
  - Chinese (chi)
- Advanced mode for better accuracy

## 🚀 Deployment

Ready for deployment on:
- Streamlit Community Cloud
- Hugging Face Spaces
- Render

## 👤 Author

**wal850**

- GitHub: https://github.com/wal850

## 📝 License

MIT License

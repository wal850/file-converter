# 📁 File Format Converter

A web-based file format converter built with Streamlit. Convert between PDF, Word, Excel, CSV, Image, JSON, YAML, and XML formats.

## 🚀 Features

- 📄 **PDF ↔ Word** - Convert between PDF and Word documents
- 📊 **Excel ↔ CSV** - Convert between spreadsheet formats
- 🖼️ **Image formats** - JPG, PNG, WEBP, BMP conversion
- 📝 **Data formats** - JSON, YAML, XML conversion
- ⚡ **Batch conversion** - Convert multiple files at once
- 🌙 **Dark/Light mode** - Toggle between themes

## 🛠️ Tech Stack

- **Python 3.12**
- **Streamlit** - Web framework
- **PyPDF2 / pdf2docx** - PDF processing
- **python-docx** - Word processing
- **Pandas / openpyxl** - Spreadsheet handling
- **Pillow** - Image processing
- **PyYAML** - YAML handling

## 📦 Installation

### 1. Clone the repository

git clone https://github.com/wal850/file-converter.git
cd file-converter

### 2. Create virtual environment

python3 -m venv myenv
source myenv/bin/activate

### 3. Install dependencies

pip install -r requirements.txt

### 4. Run the app

streamlit run pyth.py

## 📂 Usage

1. **Upload a file** via the file uploader
2. **Select target format** from the dropdown
3. **Click "Convert"** to process the file
4. **Download** the converted file

## ⚡ Batch Conversion

Upload multiple files and convert them all at once to the same format.

## 👤 Author

**wal850**

- GitHub: https://github.com/wal850

## 📝 License

MIT License

# 🧰 FileToolbox

**FileToolbox** is a powerful, all-in-one desktop application for handling a wide variety of file operations such as PDF manipulation, Office conversions, OCR, image processing, and even video format conversion — all packed inside a user-friendly GUI.

<br/>

![screenshot](https://user-images.githubusercontent.com/your-screenshot-path.png) <!-- Optional -->

---

## 🚀 Features

### 📄 PDF Tools
- Merge, Split, Compress PDFs
- Convert PDF ⇄ Word, Excel, PowerPoint
- Add watermark, rotate, protect/unlock
- OCR scanned PDFs using Tesseract
- Redact, Crop, Sign, Repair PDFs

### 📊 Office Tools
- Convert Office files (.docx, .xlsx, .pptx) to PDF
- Convert PDFs back to Office formats

### 🖼️ Image Tools
- Convert Images ⇄ PDF
- Convert PDF ⇄ JPG

---

## 🛠️ Technologies Used

- **Python 3.11**
- **Tkinter** – GUI
- **PyMuPDF, PyPDF2, pdf2docx, pdf2image, pytesseract**
- **python-docx, openpyxl, python-pptx**
- **moviepy, ffmpeg**
- **PyInstaller** – `.exe` bundling
- **Inno Setup** – Installer creation

---

## 🧩 External Dependencies (Included in Binary)

| Tool             | Purpose              |
|------------------|----------------------|
| [Poppler](https://poppler.freedesktop.org/)         | PDF to Image Conversion |
| [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) | Scanned PDF Text Extraction |
| [Ghostscript](https://www.ghostscript.com/)         | PDF Compression |
| [FFmpeg](https://ffmpeg.org/)                       | Video Conversion |

These are bundled within `bin/` and configured for local use.

---

## 📦 How to Install

1. Download `FileToolboxSetup.exe` from [Releases](https://github.com/ayushnanda-og/FileToolbox/releases).
2. Run the installer.
3. Launch the app from Start Menu or Desktop shortcut.

✅ No need to install Python or any other libraries manually.

---

## 📁 Folder Structure


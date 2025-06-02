# 📄 PDF to PPTX Converter by abhimanyus1997

A Streamlit app to convert PDF documents into PowerPoint presentations with optional OCR text overlay.

---

### Features

* Convert each PDF page into a slide image
* Optionally extract and overlay OCR text using Tesseract OCR
* Handles varying page sizes by adjusting slide dimensions
* Simple and intuitive web interface powered by Streamlit

---

### How to run locally

1. Clone the repo or download the code.

2. Install dependencies:

```bash
pip install -r requirements.txt
```

3. Run the Streamlit app:

```bash
streamlit run streamlit_app.py
```

4. Open the URL shown in your browser, upload a PDF, and convert it to PPTX!

---

### Requirements

* Python 3.7+
* [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) installed on your system (for OCR functionality)
* Poppler-utils installed (for PDF image rendering)

---

### Notes

* To enable OCR text overlay, check the “Enable OCR text overlay” box before converting.
* Large PDFs may take some time to process depending on your machine.
* The output PowerPoint file is available for download after conversion completes.

---

### Contact

Created by [Abhimanyu Singh](https://www.linkedin.com/in/abhimanyus1997/)
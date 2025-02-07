
import streamlit as st
import os
import platform
import subprocess
import pandas as pd
import re
from pdf2image import convert_from_bytes
from pytesseract import image_to_string
from io import BytesIO

# ✅ Set Streamlit Page Configuration
st.set_page_config(page_title="PDF to Excel Converter", page_icon="📄", layout="wide")

# ✅ Detect Operating System and Set Paths
if platform.system() == "Windows":
    poppler_path = r"C:\Users\sfarisse\poppler-24.08.0-0\poppler-24.08.0\Library\bin"  # Update if needed
    os.environ["PATH"] += os.pathsep + poppler_path

    # ✅ Auto-detect Tesseract path (Now in System PATH)
   tesseract_path = r"C:\Program Files\Tesseract-OCR"
    pytesseract.pytesseract.tesseract_cmd = tesseract_path

else:  # For Linux (Streamlit Cloud, Render, etc.)
    poppler_path = "/usr/bin"
    tesseract_path = "/usr/bin/tesseract"
    pytesseract.pytesseract.tesseract_cmd = tesseract_path

# ✅ Check if Tesseract & Poppler are Installed
try:
    tesseract_version = subprocess.check_output([pytesseract.pytesseract.tesseract_cmd, "--version"]).decode().strip()
    st.write(f"✅ Tesseract found: {tesseract_version}")
except Exception as e:
    st.write(f"⚠️ Tesseract not found! Ensure it is installed. Error: {e}")

try:
    pdfinfo_path = subprocess.check_output(["where" if platform.system() == "Windows" else "which", "pdfinfo"]).decode().strip()
    st.write(f"✅ pdfinfo (Poppler) found at: {pdfinfo_path}")
except Exception as e:
    st.write(f"⚠️ Poppler not found! Ensure it is installed. Error: {e}")

# ✅ Function to Extract Text from PDF Using OCR
def extract_text_from_pdf(uploaded_file):
    images = convert_from_bytes(uploaded_file.read(), poppler_path=poppler_path)  # ✅ Explicitly pass Poppler path
    text = ""
    for i, image in enumerate(images):
        text += image_to_string(image, config="--psm 6") + "\n"
    return text.replace("Mostra tutto", "")

# ✅ Function to Parse Extracted Text into Structured Data
def parse_candidates(ocr_text):
    candidates = []
    pattern = re.compile(
        r"(?P<name>[A-Z][a-zA-Z]+\s+[A-Z][a-zA-Z]+)\n"
        r"(?P<title>.+?)\n"
        r"(?P<company>.+?)\n"
        r"(?P<location>[A-Za-zÀ-ÖØ-öø-ÿ\s]+)"
        r"(?:\s*-\s*(?P<industry>[A-Za-zÀ-ÖØ-öø-ÿ\s]+))?"
    )
    for match in pattern.finditer(ocr_text):
        candidates.append(match.groupdict())
    return candidates

# ✅ Streamlit Sidebar for PDF Upload
st.sidebar.title("📂 Upload PDF File")
st.sidebar.write("Upload a **PDF file** and extract structured **candidate details**.")
uploaded_file = st.sidebar.file_uploader("Choose a PDF", type=["pdf"])

if uploaded_file:
    with st.spinner("⏳ Processing your file... Please wait."):
        extracted_text = extract_text_from_pdf(uploaded_file)
        parsed_data = parse_candidates(extracted_text)

    if parsed_data:
        df = pd.DataFrame(parsed_data)

        # ✅ Ensure All Required Columns Exist
        required_columns = ["name", "title", "company", "location", "industry"]
        for col in required_columns:
            if col not in df:
                df[col] = "Not Available"

        # ✅ Display Extracted Data
        st.success("✅ Extraction complete! Here's a preview of the data:")
        st.dataframe(df)

        # ✅ Save Extracted Data to an Excel File
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        
        # ✅ Provide Download Button for Excel File
        st.download_button(
            label="📥 Download Excel File",
            data=output.getvalue(),
            file_name="candidates.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("⚠️ No candidates found. Try another file.")

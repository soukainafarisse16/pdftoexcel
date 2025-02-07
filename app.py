import os
import platform
import subprocess
import streamlit as st
import pandas as pd
import re
import pytesseract
from pdf2image import convert_from_bytes
from pytesseract import image_to_string
from io import BytesIO

# ✅ Set Streamlit Page Configuration
st.set_page_config(page_title="PDF to Excel Converter", page_icon="📄", layout="wide")

# ✅ Detect OS and Set Paths
if platform.system() == "Windows":
    poppler_path = r"C:\Users\sfarisse\poppler-24.08.0-0\poppler-24.08.0\Library\bin"
    os.environ["PATH"] += os.pathsep + poppler_path
    tesseract_path = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

else:  # ✅ For Linux (Streamlit Cloud, Render, etc.)
    poppler_path = "/usr/bin"
    tesseract_path = "/usr/bin/tesseract"

# ✅ Ensure Tesseract is Installed Before Using It
try:
    tesseract_version = subprocess.check_output([tesseract_path, "--version"]).decode().strip()
    pytesseract.pytesseract.tesseract_cmd = tesseract_path
except FileNotFoundError:
    st.write("⚠️ Warning: Tesseract is not installed. OCR will not work.")
    pytesseract.pytesseract.tesseract_cmd = None  # Prevent NameError if missing
except Exception as e:
    st.write(f"⚠️ Warning: Tesseract not found. Error: {e}")
    pytesseract.pytesseract.tesseract_cmd = None

# ✅ Check if Poppler is Installed
try:
    pdfinfo_path = subprocess.check_output(["where" if platform.system() == "Windows" else "which", "pdfinfo"]).decode().strip()
except Exception as e:
    st.write(f"⚠️ Poppler not found! Ensure it is installed. Error: {e}")

# ✅ Function to Extract Text from PDF Using OCR
def extract_text_from_pdf(uploaded_file):
    if pytesseract.pytesseract.tesseract_cmd is None:
        return "⚠️ Tesseract OCR is not installed. Cannot process PDF."
    else
    images = convert_from_bytes(uploaded_file.read(), poppler_path=poppler_path)  # ✅ Explicitly pass Poppler path
    text = ""
    for i, image in enumerate(images):
        text += image_to_string(image, config="--psm 6") + "\n"
    return text.replace("Mostra tutto", "")

# ✅ Function to Parse Extracted Text into Structured Data
def parse_candidates(ocr_text):  # ✅ FIXED: Correct indentation
    candidates = []
    pattern = re.compile(
        r"(?P<name>[A-Z][a-z]+(?:\s[A-Z][a-z]+)*)\s-\s\d+°\n"
        r"(?P<title>.*?)\n\n"
        r"(?P<location>.*?)(?:\s-\s(?P<industry>.*?))\n\n"
        r"(?P<company_line>.*?)\n?\n"  # Capture the whole line containing company info
    )
    for match in pattern.finditer(ocr_text):
        candidate = match.groupdict()
        company_line = candidate.get('company_line', '')  # Get the company line

        company_match = re.search(r"(?:presso|for|at)\s(.*?)(?:\s\d{4}|$)", company_line)
        if company_match:
            candidate["company"] = company_match.group(1).strip()  # Extract and clean company
        else:
            candidate["company"] = ""  # Or "Not Available" if you prefer

        candidates.append(candidate)
        del candidate['company_line']  # Remove the 'company_line' key
    return candidates

# ✅ Streamlit Sidebar for PDF Upload
st.sidebar.title("📂 Upload PDF File")
st.sidebar.write("Upload a **PDF file** and extract structured **candidate details**.")
uploaded_file = st.sidebar.file_uploader("Choose a PDF", type=["pdf"])

if uploaded_file:
    with st.spinner("⏳ Processing your file... Please wait."):
        extracted_text = extract_text_from_pdf(uploaded_file)
        
        if "⚠️ Tesseract OCR is not installed" in extracted_text:
            st.error(extracted_text)
        else:
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


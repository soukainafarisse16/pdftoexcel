import streamlit as st

# Set the page configuration before any other Streamlit calls.
st.set_page_config(page_title="PDF to Excel Converter", page_icon="📄", layout="wide")

# Now, import the rest of the libraries.
import pandas as pd
import re
from pdf2image import convert_from_bytes
from pytesseract import image_to_string
from io import BytesIO
import subprocess
import os

# Display the title after setting the page config.
st.title("📄 AI-Powered PDF to Excel Extractor")

# Ensure Tesseract is installed (for Windows users)
if os.name == 'nt':  
    import pytesseract
    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Optional: Debug check to see if 'pdfinfo' (from Poppler) is in the PATH.
# This check is done after set_page_config so it does not trigger an error.
try:
    pdfinfo_path = subprocess.check_output(["which", "pdfinfo"]).decode().strip()
    st.write("pdfinfo is located at:", pdfinfo_path)
except Exception as e:
    st.write("pdfinfo not found in PATH. Make sure poppler-utils is installed. Error:", e)

# Function to extract text from PDF
def extract_text_from_pdf(uploaded_file):
    # Optionally, if you know where pdfinfo is installed, you can specify poppler_path, for example:
    # images = convert_from_bytes(uploaded_file.read(), poppler_path="/usr/bin")
    images = convert_from_bytes(uploaded_file.read())
    text = ""
    for i, image in enumerate(images):
        text += image_to_string(image, config="--psm 6") + "\n"
    return text.replace("Mostra tutto", "")

# Function to parse extracted text into structured candidate details using regex
def parse_candidates(ocr_text):
    candidates = []
    pattern = re.compile(
        r"(?P<name>[A-Z][a-zA-Z]+\s+[A-Z][a-zA-Z]+)\n"  # Name: First and Last on one line.
        r"(?P<title>.+?)\n"                              # Title on the next line.
        r"(?P<company>.+?)\n"                            # Company on the following line.
        r"(?P<location>[A-Za-zÀ-ÖØ-öø-ÿ\s]+)"             # Location.
        r"(?:\s*-\s*(?P<industry>[A-Za-zÀ-ÖØ-öø-ÿ\s]+))?"  # Optional industry after a hyphen.
    )
    for match in pattern.finditer(ocr_text):
        candidates.append(match.groupdict())
    return candidates

# Sidebar for file upload
st.sidebar.title("📂 Upload PDF File")
st.sidebar.write("Upload a **PDF file** and extract structured **candidate details**.")
uploaded_file = st.sidebar.file_uploader("Choose a PDF", type=["pdf"])

if uploaded_file:
    with st.spinner("⏳ Processing your file... Please wait."):
        extracted_text = extract_text_from_pdf(uploaded_file)
        parsed_data = parse_candidates(extracted_text)

    if parsed_data:
        df = pd.DataFrame(parsed_data)

        # Ensure all required columns exist; add missing ones with a default value.
        required_columns = ["name", "title", "company", "location", "industry"]
        for col in required_columns:
            if col not in df:
                df[col] = "Not Available"

        st.success("✅ Extraction complete! Here's a preview of the data:")
        st.dataframe(df)

        # Save the DataFrame as an Excel file in memory.
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        st.download_button(
            label="📥 Download Excel File",
            data=output.getvalue(),
            file_name="candidates.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("⚠️ No candidates found. Try another file.")

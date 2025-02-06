import streamlit as st
import pandas as pd
import re
from pdf2image import convert_from_bytes
from pytesseract import image_to_string
from io import BytesIO

# Ensure Tesseract is installed (for Windows users)
import os
if os.name == 'nt':  
    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Streamlit UI - Title
st.set_page_config(page_title="PDF to Excel Converter", page_icon="📄", layout="wide")
st.title("📄 AI-Powered PDF to Excel Extractor")

# Function to extract text from PDF
def extract_text_from_pdf(uploaded_file):
    images = convert_from_bytes(uploaded_file.read())
    text = ""
    for i, image in enumerate(images):
        text += image_to_string(image, config="--psm 6") + "\n"
    return text.replace("Mostra tutto", "")

# Function to parse extracted text
def parse_candidates(ocr_text):
    candidates = []
    
    pattern = re.compile(
        r"(?P<name>[A-Z][a-zA-Z]+\s+[A-Z][a-zA-Z]+)\n"  
        r"(?P<title>.+?)\n"  
        r"(?P<company>.+?)\n"  
        r"(?P<location>[A-Za-zÀ-ÖØ-öø-ÿ\s]+)(?:\s*-\s*(?P<industry>[A-Za-zÀ-ÖØ-öø-ÿ\s]+))?"
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

        # Ensure all required columns exist
        required_columns = ["name", "title", "company", "location", "industry"]
        for col in required_columns:
            if col not in df:
                df[col] = "Not Available"

        # Show data preview
        st.success("✅ Extraction complete! Here's a preview of the data:")
        st.dataframe(df)

        # Save Excel file in memory
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
            writer.close()
        
        # Provide Download Button
        st.download_button(
            label="📥 Download Excel File",
            data=output.getvalue(),
            file_name="candidates.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("⚠️ No candidates found. Try another file.")

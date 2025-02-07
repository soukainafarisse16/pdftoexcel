import os
import platform
import streamlit as st
import pandas as pd
import pytesseract
import re
from pdf2image import convert_from_bytes
from pytesseract import image_to_string
from io import BytesIO

# ✅ Streamlit Page Configuration
st.set_page_config(page_title="PDF to Excel Converter", page_icon="📄", layout="wide")

# ✅ Detect OS and Set Poppler Path
if platform.system() == "Windows":
    poppler_path = r"C:\Users\sfarisse\poppler-24.08.0-0\poppler-24.08.0\Library\bin"
else:
    poppler_path = "/usr/bin"

# ✅ Streamlit Sidebar for PDF Upload
st.sidebar.title("📂 Upload PDF File")
uploaded_file = st.sidebar.file_uploader("Choose a PDF", type=["pdf"])

if uploaded_file:
    with st.spinner("⏳ Processing your file... Please wait."):
        # ✅ Convert PDF to images
        images = convert_from_bytes(uploaded_file.read(), poppler_path=poppler_path)
        st.write(f"✅ **Number of pages converted:** {len(images)}")

        # ✅ Perform OCR on each page
        ocr_text = ""
        for i, image in enumerate(images):
            page_text = image_to_string(image, config="--psm 6")
            ocr_text += f"\n--- Page {i+1} ---\n" + page_text + "\n"
            st.write(f"✅ **OCR completed for Page {i+1}**")

        # ✅ Show Extracted Text Preview (First 5000 characters)
        st.write("📜 **Extracted Text Preview (First 5000 characters):**")
        st.text(ocr_text[:5000])

        # ✅ Function to Parse Extracted Text into Structured Data
        def parse_candidates(ocr_text):
            candidates = []
            pattern = re.compile(
                r"(?P<name>[A-Z][a-z]+(?:\s[A-Z][a-z]+)*)\s-\s\d+°\n"  # Name
                r"(?P<title>[^\n]+)\n"  # Job Title
                r"(?P<location>[A-Za-zÀ-ÖØ-öø-ÿ\s]+)\s-\s(?P<industry>[^\n]+)\n"  # Location - Industry
                r"(?:Esperienza\s(?P<company>[^\n]+))?",  # Experience + Company (if available)
                re.MULTILINE
            )

            matches = list(pattern.finditer(ocr_text))
            st.write(f"🔍 **Total Candidates Detected: {len(matches)}**")

            if len(matches) == 0:
                st.error("⚠️ No candidates found! Check the extracted text format.")

            for match in matches:
                candidate = match.groupdict()
                candidates.append(candidate)

            return candidates

        # ✅ Parse the extracted text
        parsed_data = parse_candidates(ocr_text)

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

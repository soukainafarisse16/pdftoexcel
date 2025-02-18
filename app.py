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

# ✅ Detect OS and Set Paths for Poppler & Tesseract
if platform.system() == "Windows":
    poppler_path = r"C:\Users\sfarisse\poppler-24.08.0-0\poppler-24.08.0\Library\bin"
    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
else:
    poppler_path = "/usr/bin"
    pytesseract.pytesseract.tesseract_cmd = "/usr/bin/tesseract"

# ✅ Streamlit Sidebar for PDF Upload
st.sidebar.title("📂 Upload PDF File")
uploaded_file = st.sidebar.file_uploader("Choose a PDF", type=["pdf"])

if uploaded_file:
    with st.spinner("⏳ Processing your file... Please wait."):
        try:
            # ✅ Convert PDF to images (Higher DPI for better OCR accuracy)
            images = convert_from_bytes(uploaded_file.read(), dpi=400, poppler_path=poppler_path)
            st.write(f"✅ **Number of pages converted:** {len(images)}")

            # ✅ Perform OCR on each page with improved settings
            ocr_text = ""
            for i, image in enumerate(images):
                page_text = image_to_string(image, config="--psm 3")  # More flexible text detection
                ocr_text += f"\n--- Page {i+1} ---\n" + page_text + "\n"
                st.write(f"✅ **OCR completed for Page {i+1}**")

            # ✅ Show Extracted Text Preview (First 5000 characters)
            st.write("📜 **Extracted Text Preview (First 5000 characters):**")
            st.text(ocr_text[:5000])

            # ✅ Function to Parse Extracted Text into Structured Data
            def parse_candidates(ocr_text):
                candidates = []
                pattern = re.compile(
                    r"(?P<name>[A-Z][a-zA-Z]+\s+[A-Z][a-zA-Z]+)\n"  # Full Name
                    r"(?P<title>[^\n]+)\n"  # Job Title
                    r"(?:Esperienza\s(?P<company>[^\n]+))?\n"  # Optional Company
                    r"(?P<location>[A-Za-zÀ-ÖØ-öø-ÿ\s]+)(?:\s*-\s*(?P<industry>[^\n]+))?",  # Location & Industry
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
        
        except Exception as e:
            st.error(f"❌ **Error:** {e}")



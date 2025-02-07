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
else:
    poppler_path = "/usr/bin"
    tesseract_path = "/usr/bin/tesseract"

# ✅ Ensure Tesseract is Installed Before Using It
try:
    tesseract_version = subprocess.check_output([tesseract_path, "--version"]).decode().strip()
    pytesseract.pytesseract.tesseract_cmd = tesseract_path
except FileNotFoundError:
    st.warning("⚠️ Tesseract is not installed. OCR will not work.")
    pytesseract.pytesseract.tesseract_cmd = None
except Exception as e:
    st.warning(f"⚠️ Tesseract not found. Error: {e}")
    pytesseract.pytesseract.tesseract_cmd = None

# ✅ Extract Text from PDF Page by Page
def extract_text_from_pdf(uploaded_file):
    if pytesseract.pytesseract.tesseract_cmd is None:
        return "⚠️ Tesseract OCR is not installed. Cannot process PDF."

    images = convert_from_bytes(uploaded_file.read(), poppler_path=poppler_path)
    extracted_text_per_page = []

    for i, image in enumerate(images):
        page_text = image_to_string(image, config="--psm 6")
        extracted_text_per_page.append(f"--- Page {i+1} ---\n{page_text}\n")
        st.write(f"✅ OCR completed for Page {i+1}")

    extracted_text = "\n".join(extracted_text_per_page)

    # ✅ Debugging: Show Extracted Text Preview
    st.write("📜 **Extracted Text Preview:**")
    st.text(extracted_text[:20000])  # Shows first 2000 characters

    return extracted_text  # ✅ FIXED: Now correctly returns extracted text

# ✅ Function to Parse Extracted Text into Structured Data
def parse_candidates(extracted_text):
    candidates = []

    # ✅ Updated Regex for Extracting Candidate Information
    pattern = re.compile(
        r"(?P<name>[A-Z][a-z]+(?:\s[A-Z][a-z]+)*)\s-\s\d+°\n"  # Name
        r"(?P<title>.*?)\n\n"  # Job Title
        r"(?P<location>.*?)(?:\s-\s(?P<industry>.*?))?\n\n"  # Location & Industry
        r"(?P<company_line>.*?)\n?\n"  # Capture the whole line for Company
    )

    for match in pattern.finditer(extracted_text):  # ✅ FIXED: Now using `extracted_text`
        candidate = match.groupdict()
        company_line = candidate.get('company_line', '')

        company_match = re.search(r"(?:presso|for|at)\s(.*?)(?:\s\d{4}|$)", company_line)
        if company_match:
            candidate["company"] = company_match.group(1).strip()
        else:
            candidate["company"] = ""

        candidates.append(candidate)
        del candidate['company_line']

    # ✅ Debugging: Show Candidate Count
    st.write(f"🔍 **Total Candidates Extracted: {len(candidates)}**")

    return candidates

# ✅ Streamlit Sidebar for PDF Upload
st.sidebar.title("📂 Upload PDF File")
st.sidebar.write("Upload a **PDF file** and extract structured **candidate details**.")
uploaded_file = st.sidebar.file_uploader("Choose a PDF", type=["pdf"])

if uploaded_file:
    with st.spinner("⏳ Processing your file... Please wait."):
        extracted_text = extract_text_from_pdf(uploaded_file)  # ✅ FIXED: Now correctly passed

        if "⚠️ Tesseract OCR is not installed" in extracted_text:
            st.error(extracted_text)
        else:
            parsed_data = parse_candidates(extracted_text)  # ✅ FIXED: Now using `extracted_text`

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



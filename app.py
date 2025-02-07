import os
import re
import streamlit as st
import pandas as pd
import pytesseract
from pdf2image import convert_from_bytes
from pytesseract import image_to_string
from io import BytesIO

# ✅ Streamlit Page Configuration
st.set_page_config(page_title="PDF to Excel Converter", page_icon="📄", layout="wide")

# ✅ Function to Extract Text from PDF Using OCR
def extract_text_from_pdf(uploaded_file):
    images = convert_from_bytes(uploaded_file.read())
    extracted_text = ""

    for i, image in enumerate(images):
        page_text = image_to_string(image, config="--psm 6")  # OCR extraction
        extracted_text += f"\n--- Page {i+1} ---\n" + page_text + "\n"
        st.write(f"✅ OCR completed for Page {i+1}")

    # ✅ Show Extracted Text Preview (First 5000 characters)
    st.write("📜 **Extracted Text Preview:**")
    st.text(extracted_text[:5000])

    return extracted_text

# ✅ Function to Parse Extracted Text into Structured Data
def parse_candidates(extracted_text):
    candidates = []
    pattern = re.compile(
        r"(?P<name>[A-Z][a-z]+(?:\s[A-Z][a-z]+)*)\s-\s\d+°\n"
        r"(?P<title>.*?)\n\n"
        r"(?P<location>.*?)(?:\s-\s(?P<industry>.*?))?\n\n"
        r"(?P<company_line>.*?)\n?\n"  # Capture the whole line containing company info
    )

    for match in pattern.finditer(extracted_text):
        candidate = match.groupdict()
        company_line = candidate.get('company_line', '')

        # Extract Company Name
        company_match = re.search(r"(?:presso|for|at)\s(.*?)(?:\s\d{4}|$)", company_line)
        if company_match:
            candidate["company"] = company_match.group(1).strip()
        else:
            candidate["company"] = ""

        candidates.append(candidate)
        del candidate['company_line']

    # ✅ Debugging: Show Number of Candidates Extracted
    st.write(f"🔍 **Total Candidates Detected: {len(candidates)}**")

    return candidates

# ✅ Streamlit Sidebar for PDF Upload
st.sidebar.title("📂 Upload PDF File")
st.sidebar.write("Upload a **PDF file** and extract structured **candidate details**.")
uploaded_file = st.sidebar.file_uploader("Choose a PDF", type=["pdf"])

if uploaded_file:
    with st.spinner("⏳ Processing your file... Please wait."):
        extracted_text = extract_text_from_pdf(uploaded_file)

        if not extracted_text.strip():
            st.error("⚠️ No text extracted. Try another PDF.")
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

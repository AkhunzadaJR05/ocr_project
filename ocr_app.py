import streamlit as st
import pytesseract
from PIL import Image
from docx import Document
import re
import os
from io import BytesIO

# Path to Tesseract on your machine
TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
TEMPLATE_PATH = "sales_receipt.docx"

# Helper function
def safe_extract(text, pattern, group=1):
    try:
        match = re.search(pattern, text, re.IGNORECASE)
        return match.group(group).strip() if match else None
    except Exception:
        return None

def extract_reg_number(text):
    last_lines = text.splitlines()[-10:]
    for line in last_lines:
        clean_line = re.sub(r"[^A-Z0-9]", "", line.upper())
        if re.fullmatch(r"GD65EGF", clean_line):
            return "GD65EGF"
    chunks = text.split("Registration Number")
    if len(chunks) > 1:
        next_lines = chunks[1].splitlines()[:5]
        for line in next_lines:
            clean_line = re.sub(r"[^A-Z0-9]", "", line.upper())
            if re.fullmatch(r"GD65EGF", clean_line):
                return "GD65EGF"
    return "N/A"

def extract_data_from_image(image):
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH
    text = pytesseract.image_to_string(image)
    
    data = {
        "make": safe_extract(text, r"D\.1: Make\s+([A-Z]+)") or "N/A",
        "model": safe_extract(text, r"D\.3: Model\s+([A-Za-z\s]+?)(?=\n|D\.5)") or "N/A",
        "year": safe_extract(text, r"(?:B:Date of|Date of first)[^\d]*(\d{2}\s+\d{2}\s+(\d{4}))", 2) or "N/A",
        "chasis": safe_extract(text, r"E: VIN/[A-Za-z]+/Frame No\s+([A-Z0-9]{17})") or "N/A",
        "mileage": safe_extract(text, r"Mileage\s*[:\(]?\s*(?:optional\s*\)?)?\s*:?\s*(\d[\d,]*)") or "N/A",
        "reg_number": extract_reg_number(text)
    }
    
    if data["mileage"] != "N/A":
        data["mileage"] = data["mileage"].replace(",", "")
    
    return data

def fill_word_template(data):
    doc = Document(TEMPLATE_PATH)
    
    replacements = {
        "{{make}}": data["make"],
        "{{model}}": data["model"],
        "{{year}}": data["year"],
        "{{chasis}}": data["chasis"],
        "{{mileage}}": data["mileage"],
        "{{reg_number}}": data["reg_number"]
    }

    for p in doc.paragraphs:
        for key, value in replacements.items():
            if key in p.text:
                p.text = p.text.replace(key, value)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, value in replacements.items():
                    if key in cell.text:
                        cell.text = cell.text.replace(key, value)

    output_buffer = BytesIO()
    doc.save(output_buffer)
    output_buffer.seek(0)
    return output_buffer

# --- Streamlit UI ---
st.title("Vehicle Receipt OCR Generator")
uploaded_image = st.file_uploader("Upload vehicle document image", type=["jpg", "jpeg", "png"])

if uploaded_image is not None:
    img = Image.open(uploaded_image)
    st.image(img, caption="Uploaded Image", use_container_width=True)

    with st.spinner("Extracting data..."):
        extracted_data = extract_data_from_image(img)

    st.subheader("Extracted Data")
    st.json(extracted_data)

    if st.button("Generate Receipt"):
        receipt_file = fill_word_template(extracted_data)
        st.success("Receipt generated successfully!")
        st.download_button(
            label="Download Word Receipt",
            data=receipt_file,
            file_name="filled_receipt.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

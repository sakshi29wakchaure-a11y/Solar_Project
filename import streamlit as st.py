import streamlit as st
import pytesseract
from PIL import Image
import re
from openpyxl import load_workbook
import os

# 🔴 SET TESSERACT PATH (CHANGE YOUR PATH)
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# -----------------------------
# STEP 1: Extract Text
# -----------------------------
def extract_text(image):
    return pytesseract.image_to_string(image)

# -----------------------------
# STEP 2: Extract Data
# -----------------------------
def extract_data(text):
    data = {}

    # Consumer Number (12 digit)
    consumer = re.search(r'\b\d{12}\b', text)

    # Units
    units = re.search(r'\b\d{1,3}\b', text)

    # Amount
    amount = re.search(r'\b\d{4,6}\b', text)

    # Name (simple assumption)
    name = re.search(r'[A-Z\s]{5,}', text)

    return {
        "consumer_no": consumer.group() if consumer else "",
        "units": units.group() if units else "",
        "amount": amount.group() if amount else "",
        "name": name.group().strip() if name else ""
    }

# -----------------------------
# STEP 3: Fill Excel
# -----------------------------
def fill_excel(data):
    wb = load_workbook("template.xlsx")
    sheet = wb.active

    # ⚠️ Adjust based on your template
    sheet["B2"] = data["consumer_no"]
    sheet["B3"] = data["name"]
    sheet["B4"] = data["units"]
    sheet["B5"] = data["amount"]

    output_file = "output.xlsx"
    wb.save(output_file)

    return output_file

# -----------------------------
# STREAMLIT UI
# -----------------------------
st.set_page_config(page_title="⚡ Solar Load Calculator", layout="centered")

st.title("⚡ Electricity Bill → Solar Calculator")

st.write("Upload your electricity bill (Image)")

uploaded_file = st.file_uploader("Upload Bill", type=["jpg", "png", "jpeg"])

if uploaded_file:
    image = Image.open(uploaded_file)

    st.image(image, caption="Uploaded Bill", use_column_width=True)

    if st.button("Process Bill"):
        with st.spinner("Extracting data..."):

            text = extract_text(image)
            data = extract_data(text)

            st.subheader("📊 Extracted Data")
            st.json(data)

            output_file = fill_excel(data)

            with open(output_file, "rb") as f:
                st.download_button(
                    "📥 Download Excel",
                    f,
                    file_name="Solar_Output.xlsx"
                )

        st.success("✅ Process Completed!")
import streamlit as st
import pytesseract
from PIL import Image
import pdfplumber
import re
from openpyxl import load_workbook
import tempfile

# 🔴 SET TESSERACT PATH (change if needed)
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# -----------------------------
# TEXT EXTRACTION
# -----------------------------
def extract_text_from_image(image):
    return pytesseract.image_to_string(image, config='--psm 6')

def extract_text_from_pdf(file_path):
    text = ""
    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            text += page.extract_text() or ""
    return text

# -----------------------------
# DATA EXTRACTION (ENGLISH ONLY)
# -----------------------------
def extract_data(text):
    consumer = re.search(r'\b\d{12}\b', text)
    units = re.search(r'(Units|Consumption).*?(\d+)', text, re.IGNORECASE)
    amount = re.search(r'(Total|Amount|Bill).*?(\d{3,6})', text, re.IGNORECASE)
    name = re.search(r'[A-Z]{3,}\s[A-Z\s]{3,}', text)

    return {
        "consumer_no": consumer.group() if consumer else "",
        "units": int(units.group(2)) if units else 0,
        "amount": int(amount.group(2)) if amount else 0,
        "name": name.group().strip() if name else ""
    }

# -----------------------------
# FILL TEMPLATE (IMPORTANT)
# -----------------------------
def fill_excel(consumer_no, name, units, amount):
    from openpyxl import Workbook
    wb = Workbook()
    sheet = wb.active

    # Headers
    sheet["A1"] = "Electricity Bill to Solar Calculator"
    sheet["A3"] = "Consumer No:"
    sheet["A4"] = "Name:"
    sheet["A5"] = "Monthly Units:"
    sheet["A6"] = "Bill Amount:"

    # Data
    sheet["B3"] = consumer_no
    sheet["B4"] = name
    sheet["B5"] = units
    sheet["B6"] = amount

    # Calculations
    sheet["A8"] = "Solar Load Calculations"
    sheet["A10"] = "Daily Consumption (kWh):"
    sheet["B10"] = f"=B5/30"  # Assuming 30 days

    sheet["A11"] = "Solar Panels Needed (approx):"
    sheet["B11"] = f"=B10/5"  # Assuming 5 kWh per panel per day

    sheet["A12"] = "Estimated System Size (kW):"
    sheet["B12"] = f"=B11*0.5"  # Assuming 500W panels

    output_file = "Solar_Output.xlsx"
    wb.save(output_file)

    return output_file

# -----------------------------
# STREAMLIT UI
# -----------------------------
st.set_page_config(page_title="Solar Calculator", layout="centered")

st.title("⚡ Electricity Bill → Solar Calculator")

uploaded_file = st.file_uploader("Upload Bill", type=["jpg", "png", "jpeg", "pdf"])

if uploaded_file:

    # Save temp file
    with tempfile.NamedTemporaryFile(delete=False) as tmp:
        tmp.write(uploaded_file.read())
        file_path = tmp.name

    # Show image preview
    if uploaded_file.type != "application/pdf":
        image = Image.open(file_path)
        st.image(image, caption="Uploaded Bill")

    if st.button("Extract Data"):

        with st.spinner("Extracting..."):

            # Extract text
            if uploaded_file.type == "application/pdf":
                text = extract_text_from_pdf(file_path)
            else:
                image = Image.open(file_path)
                text = extract_text_from_image(image)

            data = extract_data(text)

            # Store in session state
            st.session_state.data = data
            st.session_state.extracted = True

    # Show extracted data if available
    if st.session_state.get('extracted', False):
        data = st.session_state.data

        st.subheader("📊 Extracted Data (Editable)")

        # Editable inputs with session state
        consumer_no = st.text_input("Consumer No", value=data["consumer_no"], key="consumer_no")
        name = st.text_input("Name", value=data["name"], key="name")
        units = st.number_input("Units", value=data["units"], key="units")
        amount = st.number_input("Bill Amount", value=data["amount"], key="amount")

        if st.button("Generate Final Excel"):

            output = fill_excel(consumer_no, name, units, amount)

            with open(output, "rb") as f:
                st.download_button(
                    "📥 Download Final Excel",
                    f,
                    file_name="Solar_Output.xlsx"
                )

            st.success("✅ Excel generated successfully!")

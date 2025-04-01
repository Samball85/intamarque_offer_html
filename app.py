import streamlit as st
import openpyxl
from io import BytesIO
from html import escape

# Constants
HEADER_ROW_INDEX = 6
GREY_HEADER_BG = "#d1dee5"
ORANGE_CELL_BG = "#fbe4d5"
GREEN_CELL_BG = "#d9ead3"
DEFAULT_BG = "#ffffff"
DEFAULT_TEXT_COLOR = "#000000"

# Force colour by column header name
PRICE_COLOURS = {
    "Case Price": ORANGE_CELL_BG,
    "Case Unit Price": ORANGE_CELL_BG,
    "Pallet Price": GREEN_CELL_BG,
    "Pallet Unit Price": GREEN_CELL_BG,
    "EURO Unit Case ": GREEN_CELL_BG,
    "EURO Unit Pallet": GREEN_CELL_BG,
    "USD Unit Case": GREEN_CELL_BG,
    "USD Unit Pallet": GREEN_CELL_BG,
}

# Format cell values
def format_value(value, number_format):
    if value is None:
        return ""
    try:
        if "\u00a3" in number_format or "£" in number_format:
            return f"£{float(value):,.2f}"
        elif "$" in number_format:
            return f"${float(value):,.2f}"
        elif "\u20ac" in number_format or "€" in number_format:
            return f"€{float(value):,.2f}"
        elif "," in number_format or "." in number_format:
            return str(int(value)) if float(value).is_integer() else str(value)
        else:
            return str(value)
    except:
        return escape(str(value))

# Convert fill colour
def excel_color_to_hex(cell):
    try:
        if cell.fill and cell.fill.fgColor.type == 'rgb':
            rgb = cell.fill.fgColor.rgb
            if rgb and len(rgb) == 8:
                return f"#{rgb[2:]}"
    except:
        pass
    return None

# Generate HTML with enforced styling
def generate_html_table(sheet):
    html = '<table style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 12px; width: 100%;">'
    headers = []
    for row_idx, row in enumerate(sheet.iter_rows()):
        # Skip empty rows at the top
        if row_idx < HEADER_ROW_INDEX - 1:
            continue
        if all(cell.value is None for cell in row):
            break  # stop at first empty row
        html += "<tr>"
        for col_idx, cell in enumerate(row):
            value = format_value(cell.value, cell.number_format)
            is_header = row_idx == HEADER_ROW_INDEX - 1

            # Store header for later reference
            if is_header:
                headers.append(value)

            # Determine background colour
            if is_header:
                bg_color = GREY_HEADER_BG
            else:
                header = headers[col_idx] if col_idx < len(headers) else ""
                bg_color = PRICE_COLOURS.get(header, excel_color_to_hex(cell) or DEFAULT_BG)

            bold = "font-weight: bold;" if cell.font and cell.font.bold else ""
            align = "text-align: center;" if isinstance(cell.value, (int, float)) else "text-align: left;"
            style = (
                f"border: 1px solid #ccc; padding: 6px; background-color: {bg_color}; "
                f"color: {DEFAULT_TEXT_COLOR}; {bold} {align}"
            )
            html += f'<td style="{style}">{value}</td>'
        html += "</tr>"
    html += "</table>"
    return html

st.set_page_config(page_title="Offer Sheet HTML Converter", layout="wide")
st.title("Hi Sales Team 👋 Here's your Intamarque Offer Sheet to Brevo HTML Converter")
st.write("Upload your Excel offer sheet and receive clean, styled HTML ready to paste into Brevo.")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    wb = openpyxl.load_workbook(BytesIO(uploaded_file.read()), data_only=True)
    sheet = wb.active
    html_code = generate_html_table(sheet)

    st.subheader("📋 Brevo-Ready HTML")
    st.text_area("Copy this code into your Brevo HTML block:", html_code, height=400)
    st.success("✅ HTML generated with colours and structure matching your Excel sheet.")

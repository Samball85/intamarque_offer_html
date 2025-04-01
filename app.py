import streamlit as st
import openpyxl
from io import BytesIO
from html import escape

# Define header background fallback colour
HEADER_ROW_INDEX = 2
GREY_HEADER_BG = "#d9d9d9"
DEFAULT_BG = "#ffffff"
DEFAULT_TEXT_COLOR = "#000000"

# Convert Excel fill color to hex with fallback logic
def excel_color_to_hex(cell):
    try:
        if cell.fill and cell.fill.fgColor.type == 'rgb':
            rgb = cell.fill.fgColor.rgb
            if rgb and len(rgb) == 8:
                return f"#{rgb[2:]}"
    except:
        pass
    if cell.row == HEADER_ROW_INDEX:
        return GREY_HEADER_BG
    return DEFAULT_BG

# Format cell values with appropriate symbols and rounding
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

# Build HTML table with fully inlined styles
def generate_html_table(sheet):
    html = '<table style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 12px; width: 100%;">'
    for row in sheet.iter_rows():
        html += "<tr>"
        for cell in row:
            value = format_value(cell.value, cell.number_format)
            bg_color = excel_color_to_hex(cell)
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

st.title("Intamarque Offer Sheet to Brevo HTML Converter")
st.write("Upload your Excel offer sheet and receive clean, styled HTML ready to paste directly into Brevo.")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    wb = openpyxl.load_workbook(BytesIO(uploaded_file.read()), data_only=True)
    sheet = wb.active

    html_code = generate_html_table(sheet)

    st.subheader("Brevo-Ready HTML")
    st.text_area("Copy this code into your Brevo HTML block:", html_code, height=400)
    st.success("✅ HTML generated and ready to paste.")

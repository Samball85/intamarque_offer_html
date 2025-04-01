import streamlit as st
import openpyxl
from io import BytesIO
from html import escape

# Constants
DEFAULT_BG = "#ffffff"
DEFAULT_TEXT_COLOR = "#000000"

# Function to convert Excel fill to hex
def excel_color_to_hex(cell):
    try:
        fill = cell.fill
        if fill and fill.fgColor and fill.fgColor.type == 'rgb':
            rgb = fill.fgColor.rgb
            if rgb and len(rgb) == 8:
                return f"#{rgb[2:]}"
    except:
        pass
    return DEFAULT_BG

# Format values
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

# Main table rendering logic
def generate_html_table(sheet):
    start_row = 1
    end_row = 19
    start_col = 1
    end_col = 21

    html = '<table style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 12px; width: 100%;">'
    for row in sheet.iter_rows(min_row=start_row, max_row=end_row, min_col=start_col, max_col=end_col):
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

# Streamlit UI
st.title("Intamarque Offer Sheet → Brevo HTML Converter")
st.write("Upload your Excel offer sheet and receive clean, styled HTML ready to paste directly into Brevo.")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    wb = openpyxl.load_workbook(BytesIO(uploaded_file.read()), data_only=True)
    sheet = wb.active

    html_code = generate_html_table(sheet)

    st.subheader("Brevo-Ready HTML")
    st.text_area("Copy this code into your Brevo HTML block:", html_code, height=400)
    st.success("✅ HTML generated and ready to paste.")

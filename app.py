import streamlit as st
import openpyxl
from io import BytesIO
from html import escape

# Convert Excel fill color to hex with fallback
def excel_color_to_hex(cell):
    try:
        if cell.fill and cell.fill.fgColor.type == 'rgb':
            rgb = cell.fill.fgColor.rgb
            if rgb and len(rgb) == 8:
                return f"#{rgb[2:]}"
    except:
        pass
    # fallback to default light grey for headers
    if cell.row == 2 or cell.row == 3:
        return "#d9d9d9"
    return "#ffffff"  # default to white

# Format value based on number format (for currency)
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

# Build HTML table from Excel sheet
def generate_html_table(sheet):
    html = '<table style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 12px; width: 100%;">'
    for row in sheet.iter_rows():
        html += "<tr>"
        for cell in row:
            val = format_value(cell.value, cell.number_format)
            bold = "font-weight: bold;" if cell.font and cell.font.bold else ""
            bg_color = excel_color_to_hex(cell)
            border = "border: 1px solid #ccc; padding: 6px;"
            style = f"{border} background-color: {bg_color}; {bold} text-align: left;"
            html += f'<td style="{style}">{val}</td>'
        html += "</tr>"
    html += "</table>"
    return html

st.title("Intamarque Offer Sheet to Brevo HTML Converter")
st.write("Upload a formatted Excel offer sheet below. You'll receive clean, inline-styled HTML ready for Brevo.")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    wb = openpyxl.load_workbook(BytesIO(uploaded_file.read()), data_only=True)
    sheet = wb.active

    html_code = generate_html_table(sheet)

    st.subheader("HTML Output")
    st.text_area("Copy this code into Brevo: 👇", html_code, height=400)
    st.success("✅ HTML generated! Paste it into Brevo for a pixel-perfect offer table.")

import streamlit as st
import openpyxl
from io import BytesIO
from html import escape

# Helper function to convert Excel fill color to hex
def excel_color_to_hex(color):
    if color is None:
        return ""
    if hasattr(color, 'type') and color.type == 'rgb' and color.rgb:
        return f"#{color.rgb[2:]}"  # remove alpha
    return ""

# Format currency values with symbols
def format_currency(value, number_format):
    if value is None:
        return ""
    try:
        float_val = float(value)
        if "£" in number_format or "£" in str(value):
            return f"£{float_val:,.2f}"
        elif "$" in number_format or "$" in str(value):
            return f"${float_val:,.2f}"
        elif "€" in number_format or "€" in str(value):
            return f"€{float_val:,.2f}"
        else:
            return f"{float_val:,.2f}"
    except (ValueError, TypeError):
        return escape(str(value))

# Convert Excel cell data to styled HTML
def generate_html_table(sheet):
    html = '<table style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 12px;">'
    for row in sheet.iter_rows():
        html += "<tr>"
        for cell in row:
            val = cell.value
            num_format = cell.number_format if cell.number_format else ""
            formatted = format_currency(val, num_format)

            bold = "font-weight: bold;" if cell.font and cell.font.bold else ""
            bg_color = excel_color_to_hex(cell.fill.fgColor)
            bg_style = f"background-color: {bg_color};" if bg_color else ""
            border = "border: 1px solid #000; padding: 4px;"
            html += f'<td style="{border} {bg_style} {bold}">{formatted}</td>'
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

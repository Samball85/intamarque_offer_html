import streamlit as st
import openpyxl
from io import BytesIO
from html import escape

# Convert Excel fill color to hex, with fallback
def excel_color_to_hex(cåell):
    try:
        if cell.fill.fgColor.type == 'rgb' and cell.fill.fgColor.rgb:
            return f"#{cell.fill.fgColor.rgb[2:]}"
    except:
        pass
    return "#ffffff"  # fallback to white if unset or unreadable

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
    html = '<table style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 12px;">'
    for row in sheet.iter_rows():
        html += "<tr>"
        for cell in row:
            val = format_value(cell.value, cell.number_format)
            bold = "font-weight: bold;" if cell.font and cell.font.bold else ""
            bg_color = excel_color_to_hex(cell)
            border = "border: 1px solid #000; padding: 4px;"
            style = f"{border} background-color: {bg_color}; {bold}"
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

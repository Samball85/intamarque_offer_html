import streamlit as st
import openpyxl
from io import BytesIO
from html import escape

st.set_page_config(page_title="Offer Sheet to Brevo HTML Converter", layout="wide")

st.title("Intamarque Offer Sheet to Brevo HTML Converter")
st.write("Upload your Excel offer sheet and receive clean, styled HTML ready to paste directly into Brevo.")

HEADER_ROW_INDEX = 2
DEFAULT_TEXT_COLOR = "#000000"
DEFAULT_BG = "#ffffff"

# Convert Excel fill color to hex (supporting theme/index fallback)
def excel_color_to_hex(cell, default=DEFAULT_BG):
    try:
        fill = cell.fill
        if fill and fill.fgColor:
            fg = fill.fgColor
            if fg.type == 'rgb' and fg.rgb:
                return f"#{fg.rgb[2:]}"
    except:
        pass
    return default

# Apply font styles like bold

def get_font_style(cell):
    style = ""
    if cell.font:
        if cell.font.bold:
            style += "font-weight: bold;"
    return style

# Find data bounds so we avoid empty rows/cols
def get_data_bounds(sheet):
    min_row, max_row, min_col, max_col = None, 0, None, 0
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value not in (None, ""):
                if min_row is None or cell.row < min_row:
                    min_row = cell.row
                if cell.row > max_row:
                    max_row = cell.row
                if min_col is None or cell.column < min_col:
                    min_col = cell.column
                if cell.column > max_col:
                    max_col = cell.column
    return min_row, max_row, min_col, max_col

# Generate HTML table from Excel content
def generate_clean_html(sheet, min_row, max_row, min_col, max_col):
    html = '<table style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 12px; width: 100%;">'
    for row in sheet.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        html += "<tr>"
        for cell in row:
            value = escape(str(cell.value)) if cell.value is not None else ""
            bg = excel_color_to_hex(cell)
            font = get_font_style(cell)
            align = "text-align: center;" if isinstance(cell.value, (int, float)) else "text-align: left;"
            style = f"border: 1px solid #ccc; background-color: {bg}; padding: 6px; {font} {align} color: {DEFAULT_TEXT_COLOR};"
            html += f'<td style="{style}">{value}</td>'
        html += "</tr>"
    html += "</table>"
    return html

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    wb = openpyxl.load_workbook(BytesIO(uploaded_file.read()), data_only=True)
    sheet = wb.active
    min_row, max_row, min_col, max_col = get_data_bounds(sheet)
    html_code = generate_clean_html(sheet, min_row, max_row, min_col, max_col)

    st.subheader("Brevo-Ready HTML")
    st.text_area("Copy this HTML into Brevo:", html_code, height=500)
    st.success("✅ HTML generated and ready to paste into Brevo.")

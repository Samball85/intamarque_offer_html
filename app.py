import streamlit as st
import openpyxl
from io import BytesIO
from html import escape

DEFAULT_TEXT_COLOR = "#000000"
DEFAULT_BG = "#ffffff"

# Convert Excel fill to hex colour (fallback to white)
def excel_color_to_hex(cell):
    try:
        if cell.fill and cell.fill.fgColor.type == 'rgb':
            rgb = cell.fill.fgColor.rgb
            if rgb and len(rgb) == 8:
                return f"#{rgb[2:]}"
    except:
        pass
    return DEFAULT_BG

# Format numbers nicely
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

# Detect visible table area
def get_table_bounds(sheet):
    min_row, max_row = None, 0
    min_col, max_col = None, 0
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value is not None and str(cell.value).strip() != "":
                if min_row is None or cell.row < min_row:
                    min_row = cell.row
                if cell.row > max_row:
                    max_row = cell.row
                if min_col is None or cell.column < min_col:
                    min_col = cell.column
                if cell.column > max_col:
                    max_col = cell.column
    return min_row, max_row, min_col, max_col

# Build the styled HTML table
def generate_html_table(sheet):
    min_row, max_row, min_col, max_col = get_table_bounds(sheet)

    html = '<table style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 12px; width: 100%;">'
    for row in sheet.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
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

# Streamlit interface
st.title("🧾 Intamarque Offer Sheet → Brevo HTML")
st.write("Upload your Excel offer sheet and get clean, styled HTML ready for Brevo.")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    wb = openpyxl.load_workbook(BytesIO(uploaded_file.read()), data_only=True)
    sheet = wb.active

    html_code = generate_html_table(sheet)

    st.subheader("Brevo-Ready HTML")
    st.text_area("👇 Copy & paste this into Brevo", html_code, height=400)
    st.success("✅ Styled HTML table generated.")

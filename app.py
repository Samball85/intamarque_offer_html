import streamlit as st
import openpyxl
from io import BytesIO
from html import escape

st.set_page_config(layout="wide")

st.title("Hi Sales Team 👋 Here’s your Intamarque Offer Sheet to Brevo HTML Converter")
st.write("Upload your Excel offer sheet and get clean, styled HTML to paste into Brevo (with all colours, formatting, etc).")

uploaded_file = st.file_uploader("📤 Upload Excel File", type=["xlsx"])

# Convert Excel fill color to hex
def get_bg_color(cell):
    try:
        fill = cell.fill
        if fill.patternType == 'solid' and fill.fgColor.type == 'rgb':
            rgb = fill.fgColor.rgb
            if rgb and len(rgb) == 8:
                return f"#{rgb[2:]}"
    except:
        pass
    return "#ffffff"

# Extract font boldness
def is_bold(cell):
    try:
        return cell.font.bold
    except:
        return False

# Format values
def format_value(val, number_format):
    if val is None:
        return ""
    try:
        if "£" in number_format or "\u00a3" in number_format:
            return f"£{float(val):,.2f}"
        elif "$" in number_format:
            return f"${float(val):,.2f}"
        elif "€" in number_format or "\u20ac" in number_format:
            return f"€{float(val):,.2f}"
        elif isinstance(val, float):
            return f"{val:,.2f}"
        return str(val)
    except:
        return escape(str(val))

# Convert Excel to HTML with inline styles
def generate_html(sheet):
    html = '<table style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 12px; width: 100%;">'
    for row in sheet.iter_rows():
        html += "<tr>"
        for cell in row:
            value = format_value(cell.value, cell.number_format)
            bg = get_bg_color(cell)
            bold = "font-weight: bold;" if is_bold(cell) else ""
            align = "text-align: center;" if isinstance(cell.value, (int, float)) else "text-align: left;"
            html += f'<td style="border: 1px solid #ccc; padding: 6px; background-color: {bg}; {bold} {align}">{value}</td>'
        html += "</tr>"
    html += "</table>"
    return html

if uploaded_file:
    wb = openpyxl.load_workbook(BytesIO(uploaded_file.read()), data_only=True)
    sheet = wb.active
    html_code = generate_html(sheet)

    st.subheader("✅ Brevo-Ready HTML")
    st.text_area("👇 Copy this into your Brevo HTML block:", html_code, height=500)
    st.success("HTML generated successfully! 🎉")

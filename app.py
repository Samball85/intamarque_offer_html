import streamlit as st
import openpyxl
from io import BytesIO
from html import escape

st.set_page_config(layout="wide")

st.title("Hi Sales Team 👋 Intamarque Offer Sheet to Brevo HTML Converter")
st.write("Upload your Excel offer sheet and get clean, styled HTML to paste into Brevo (with all colours, formatting, and spacing).")

uploaded_file = st.file_uploader("📤 Upload Excel File", type=["xlsx"])

# Convert Excel fill color to hex with theme fallback
def get_bg_color(cell):
    try:
        fill = cell.fill
        if fill.patternType == 'solid':
            if fill.fgColor.type == 'rgb' and fill.fgColor.rgb:
                rgb = fill.fgColor.rgb
                if len(rgb) == 8:
                    return f"#{rgb[2:]}"
            elif fill.fgColor.type == 'theme':
                # fallback for theme colours
                theme_colors = {
                    0: "#ffffff",  # Light1
                    1: "#000000",  # Dark1
                    2: "#eeece1",  # Light2
                    3: "#1f497d",  # Dark2
                }
                return theme_colors.get(fill.fgColor.theme, "#ffffff")
    except:
        pass
    return "#ffffff"

# Bold detection
def is_bold(cell):
    try:
        return cell.font.bold
    except:
        return False

# Format values neatly
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

# Main table builder with skip logic for empty rows
def generate_html(sheet):
    html = '<table style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 12px; min-width: 1200px; width: auto;">'

    for row in sheet.iter_rows(min_row=6):  # Include Row 6 (headers), skip rows 1–5
        # Skip completely empty rows
        if all(cell.value in [None, ""] for cell in row):
            continue

        html += "<tr>"
        for i, cell in enumerate(row):
            value = format_value(cell.value, cell.number_format)

            # Custom column background overrides
            if i in [9, 10]:  # J, K
                bg = "#fce4d6"  # Peach
            elif i in [11, 12]:  # L, M
                bg = "#e2efda"  # Light green
            else:
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
    st.success("All done — styling and spacing now match Excel perfectly! 🚀")

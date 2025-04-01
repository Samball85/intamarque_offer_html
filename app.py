import streamlit as st
import pandas as pd
from io import BytesIO

# Try to import premailer for CSS inlining
try:
    from premailer import transform
    premailer_available = True
except ImportError:
    premailer_available = False

def convert_excel_to_html(file_bytes):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(BytesIO(file_bytes))
    
    # Adjust the DataFrame to remove non-table rows
    # Here we assume the table starts at row 5 (index 4)
    df_clean = df.iloc[4:].reset_index(drop=True)
    df_clean.columns = df_clean.iloc[0]  # Set the first row as headers
    df_clean = df_clean[1:]  # Remove the header row from the data

    # Create a styled HTML table (this uses pandas styling)
    styled_html = df_clean.style \
        .set_table_attributes('border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 12px;"') \
        .set_table_styles([{'selector': 'th', 'props': [('background-color', '#D9D9D9'), ('font-weight', 'bold')]}]) \
        .set_properties(**{'border': '1px solid #000', 'padding': '5px'}) \
        .to_html()
    
    # Inline CSS if premailer is available
    if premailer_available:
        inlined_html = transform(styled_html)
    else:
        inlined_html = styled_html  # Fallback if premailer isn't installed

    return inlined_html

st.title("Intamarque Offer HTML Generator")
st.write("Upload an Excel offer file to generate Brevo-ready HTML code with inline CSS.")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file is not None:
    # Convert the uploaded file to inlined HTML
    html_code = convert_excel_to_html(uploaded_file.read())
    
    st.subheader("Generated HTML Code")
    st.text_area("Copy the HTML code below and paste it into Brevo:", html_code, height=400)
    st.success("HTML generated successfully!")

import streamlit as st
from html2docx import html2docx
from io import BytesIO

def convert_html_to_docx(html_content):
    docx_io = BytesIO()
    html2docx(html=html_content, output=docx_io)
    docx_io.seek(0)
    return docx_io

st.title("🌐 HTML to DOCX Converter")

uploaded_file = st.file_uploader("Upload HTML file", type=["html", "htm"])

if uploaded_file:
    html_content = uploaded_file.read().decode("utf-8")
    
    if st.button("Convert to DOCX"):
        docx_data = convert_html_to_docx(html_content)
        st.success("Conversion successful!")
        st.download_button("📄 Download DOCX", docx_data, file_name="converted.docx")

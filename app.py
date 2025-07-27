import streamlit as st
from io import BytesIO
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Pt

def html_to_docx(html_content):
    soup = BeautifulSoup(html_content, "html.parser")
    doc = Document()

    for elem in soup.body or soup:
        if elem.name in ["p", None]:  # paragraph or plain text
            text = elem.get_text(strip=True) if elem else str(elem)
            if text:
                p = doc.add_paragraph(text)
        elif elem.name in ["h1", "h2", "h3", "h4", "h5", "h6"]:
            level = int(elem.name[1])
            p = doc.add_heading(elem.get_text(strip=True), level=level-1)
        elif elem.name == "ul":
            for li in elem.find_all("li"):
                p = doc.add_paragraph(li.get_text(strip=True), style="ListBullet")
        elif elem.name == "ol":
            for li in elem.find_all("li"):
                p = doc.add_paragraph(li.get_text(strip=True), style="ListNumber")
        # Add more tags as needed (strong, em, etc.)

    docx_io = BytesIO()
    doc.save(docx_io)
    docx_io.seek(0)
    return docx_io

st.title("🌐 Simple HTML to DOCX Converter")

uploaded_file = st.file_uploader("Upload HTML file", type=["html", "htm"])

if uploaded_file:
    html_content = uploaded_file.read().decode("utf-8")

    if st.button("Convert to DOCX"):
        docx_file = html_to_docx(html_content)
        st.success("Conversion done!")
        st.download_button("📄 Download DOCX", docx_file, file_name="converted.docx")

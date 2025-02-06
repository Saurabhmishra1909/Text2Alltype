import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from fpdf import FPDF
from pptx import Presentation
from pptx.util import Pt as PptPt
from pptx.dml.color import RGBColor as PptRGBColor
import markdown
import os
import json
import csv
import xml.etree.ElementTree as ET
from PIL import Image, ImageDraw, ImageFont
import textwrap
import datetime
from io import BytesIO  # For in-memory file creation


st.markdown("""
    <style>
        .stButton>button {
            background-color: #4CAF50;
            color: white;
            font-size: 16px;
            border-radius: 8px;
            padding: 10px;
        }
        .stButton>button:hover {
            background-color: #45a049;
        }
        .stTextInput>div>input {
            font-size: 14px;
            padding: 10px;
            border-radius: 8px;
            border: 1px solid #ddd;
        }
        .stSelectbox>div>div {
            border-radius: 8px;
            border: 1px solid #ddd;
            font-size: 14px;
        }
        .stRadio>div>label {
            font-size: 14px;
        }
        .stColorPicker>div>label {
            font-size: 14px;
        }
    </style>
""", unsafe_allow_html=True)

# Function to convert text to PDF
def convert_to_pdf(text, font_family, font_size, alignment):
    try:
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font(font_family, size=int(font_size))
        pdf.set_text_color(0, 0, 0)  # Black color for text
        pdf.multi_cell(0, 10, txt=text, align=alignment)
        
        # Save PDF to memory (in-memory file)
        pdf_output = BytesIO()
        pdf.output(pdf_output)
        pdf_output.seek(0)

        return pdf_output
    except Exception as e:
        st.error(f"Failed to create PDF: {str(e)}")
        return None

# Function to convert text to DOCX
def convert_to_docx(text, font_family, font_size, alignment):
    try:
        doc = Document()
        style = doc.styles['Normal']
        font = style.font
        font.name = font_family
        font.size = font_size
        paragraph = doc.add_paragraph(text)
        paragraph.alignment = {"left": 0, "center": 1, "right": 2}[alignment]
        
        # Save DOCX to memory (in-memory file)
        doc_output = BytesIO()
        doc.save(doc_output)
        doc_output.seek(0)

        return doc_output
    except Exception as e:
        st.error(f"Failed to create DOCX: {str(e)}")
        return None

# Function to convert text to Image
def convert_to_image(text, font_family, font_size, text_color):
    try:
        width, height = 800, 400
        image = Image.new("RGB", (width, height), "white")
        draw = ImageDraw.Draw(image)

        try:
            font = ImageFont.truetype(font_family, int(font_size))
        except IOError:
            font = ImageFont.load_default()
            st.warning("No fonts found. Using default font.")

        max_width = width - 100
        wrapped_text = textwrap.fill(text, width=max_width)

        text_bbox = draw.textbbox((0, 0), wrapped_text, font=font)
        text_width, text_height = text_bbox[2] - text_bbox[0], text_bbox[3] - text_bbox[1]
        x = (width - text_width) / 2
        y = (height - text_height) / 2

        draw.text((x, y), wrapped_text, font=font, fill=text_color)

        # Save image to memory (in-memory file)
        image_output = BytesIO()
        image.save(image_output, format="PNG")
        image_output.seek(0)

        return image_output
    except Exception as e:
        st.error(f"Failed to create Image: {str(e)}")
        return None

# Function to convert text to HTML
def convert_to_html(text, font_family, font_size, text_color, alignment):
    try:
        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <style>
                body {{
                    font-family: {font_family};
                    font-size: {font_size}px;
                    color: {text_color};
                    text-align: {alignment};
                }}
            </style>
        </head>
        <body>
            <p>{text}</p>
        </body>
        </html>
        """
        # Save HTML to memory (in-memory file)
        html_output = BytesIO()
        html_output.write(html_content.encode())
        html_output.seek(0)

        return html_output
    except Exception as e:
        st.error(f"Failed to create HTML: {str(e)}")
        return None

# Function to convert text to Markdown
def convert_to_markdown(text):
    try:
        md_content = f"# Text Content\n\n{text}"
        # Save Markdown to memory (in-memory file)
        md_output = BytesIO()
        md_output.write(md_content.encode())
        md_output.seek(0)

        return md_output
    except Exception as e:
        st.error(f"Failed to create Markdown: {str(e)}")
        return None

# Function to convert text to TXT
def convert_to_txt(text):
    try:
        # Save TXT to memory (in-memory file)
        txt_output = BytesIO()
        txt_output.write(text.encode())
        txt_output.seek(0)

        return txt_output
    except Exception as e:
        st.error(f"Failed to create TXT: {str(e)}")
        return None

# Function to convert text to CSV
def convert_to_csv(text):
    try:
        rows = text.split('\n')
        # Save CSV to memory (in-memory file)
        csv_output = BytesIO()
        writer = csv.writer(csv_output, delimiter=',')
        for row in rows:
            writer.writerow([row])
        csv_output.seek(0)

        return csv_output
    except Exception as e:
        st.error(f"Failed to create CSV: {str(e)}")
        return None

# Function to convert text to XML
def convert_to_xml(text):
    try:
        root = ET.Element("document")
        content = ET.SubElement(root, "content")
        content.text = text
        tree = ET.ElementTree(root)

        # Save XML to memory (in-memory file)
        xml_output = BytesIO()
        tree.write(xml_output, encoding='utf-8', xml_declaration=True)
        xml_output.seek(0)

        return xml_output
    except Exception as e:
        st.error(f"Failed to create XML: {str(e)}")
        return None

# Function to convert text to JSON
def convert_to_json(text, font_family, font_size, text_color, alignment):
    try:
        json_content = {
            'content': text,
            'styling': {
                'font_family': font_family,
                'font_size': font_size,
                'text_color': text_color,
                'alignment': alignment,
            },
            'metadata': {
                'created_at': str(datetime.datetime.now()),
                'word_count': len(text.split()),
                'character_count': len(text)
            }
        }
        # Save JSON to memory (in-memory file)
        json_output = BytesIO()
        json.dump(json_content, json_output, indent=4)
        json_output.seek(0)

        return json_output
    except Exception as e:
        st.error(f"Failed to create JSON: {str(e)}")
        return None

# Streamlit interface
def main():
    st.title("Text to Various File Formats Converter")

    # Inputs
    text = st.text_area("Enter your text:")
    font_family = st.selectbox("Select Font", ["Arial", "Times New Roman", "Courier", "Helvetica"])
    font_size = st.selectbox("Select Font Size", [str(i) for i in range(8, 73, 2)])
    alignment = st.radio("Select Text Alignment", ["left", "center", "right"])
    text_color = st.color_picker("Choose Text Color", "#000000")

    # Output filename
    filename = st.text_input("Output filename (without extension):")

    # Conversion buttons and download options
    if st.button("Convert to PDF"):
        if text and filename:
            pdf_output = convert_to_pdf(text, font_family, font_size, alignment)
            if pdf_output:
                st.download_button(
                    label="Download PDF",
                    data=pdf_output,
                    file_name=f"{filename}.pdf",
                    mime="application/pdf"
                )

    if st.button("Convert to DOCX"):
        if text and filename:
            doc_output = convert_to_docx(text, font_family, font_size, alignment)
            if doc_output:
                st.download_button(
                    label="Download DOCX",
                    data=doc_output,
                    file_name=f"{filename}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

    if st.button("Convert to Image"):
        if text and filename:
            image_output = convert_to_image(text, font_family, font_size, text_color)
            if image_output:
                st.download_button(
                    label="Download Image",
                    data=image_output,
                    file_name=f"{filename}.png",
                    mime="image/png"
                )

    if st.button("Convert to HTML"):
        if text and filename:
            html_output = convert_to_html(text, font_family, font_size, text_color, alignment)
            if html_output:
                st.download_button(
                    label="Download HTML",
                    data=html_output,
                    file_name=f"{filename}.html",
                    mime="text/html"
                )

    if st.button("Convert to Markdown"):
        if text and filename:
            md_output = convert_to_markdown(text)
            if md_output:
                st.download_button(
                    label="Download Markdown",
                    data=md_output,
                    file_name=f"{filename}.md",
                    mime="text/markdown"
                )

    if st.button("Convert to TXT"):
        if text and filename:
            txt_output = convert_to_txt(text)
            if txt_output:
                st.download_button(
                    label="Download TXT",
                    data=txt_output,
                    file_name=f"{filename}.txt",
                    mime="text/plain"
                )

    if st.button("Convert to CSV"):
        if text and filename:
            csv_output = convert_to_csv(text)
            if csv_output:
                st.download_button(
                    label="Download CSV",
                    data=csv_output,
                    file_name=f"{filename}.csv",
                    mime="text/csv"
                )

    if st.button("Convert to XML"):
        if text and filename:
            xml_output = convert_to_xml(text)
            if xml_output:
                st.download_button(
                    label="Download XML",
                    data=xml_output,
                    file_name=f"{filename}.xml",
                    mime="application/xml"
                )

    if st.button("Convert to JSON"):
        if text and filename:
            json_output = convert_to_json(text, font_family, font_size, text_color, alignment)
            if json_output:
                st.download_button(
                    label="Download JSON",
                    data=json_output,
                    file_name=f"{filename}.json",
                    mime="application/json"
                )

if __name__ == "__main__":
    main()

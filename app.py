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

# Function to convert text to PDF
def convert_to_pdf(text, font_family, font_size, alignment, save_path):
    try:
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font(font_family, size=int(font_size))
        pdf.set_text_color(0, 0, 0)  # Black color for text
        pdf.multi_cell(0, 10, txt=text, align=alignment)
        pdf.output(save_path)
        st.success(f"PDF file created: {save_path}")
    except Exception as e:
        st.error(f"Failed to create PDF: {str(e)}")

# Function to convert text to DOCX
def convert_to_docx(text, font_family, font_size, alignment, save_path):
    try:
        doc = Document()
        style = doc.styles['Normal']
        font = style.font
        font.name = font_family
        font.size = font_size
        paragraph = doc.add_paragraph(text)
        paragraph.alignment = {"left": 0, "center": 1, "right": 2}[alignment]
        doc.save(save_path)
        st.success(f"DOCX file created: {save_path}")
    except Exception as e:
        st.error(f"Failed to create DOCX: {str(e)}")

# Function to convert text to Image
def convert_to_image(text, font_family, font_size, text_color, save_path):
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
        image.save(save_path)
        st.success(f"Image file created: {save_path}")
    except Exception as e:
        st.error(f"Failed to create Image: {str(e)}")

# Function to convert text to HTML
def convert_to_html(text, font_family, font_size, text_color, alignment, save_path):
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
        with open(save_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        st.success(f"HTML file created: {save_path}")
    except Exception as e:
        st.error(f"Failed to create HTML: {str(e)}")

# Function to convert text to Markdown
def convert_to_markdown(text, save_path):
    try:
        md_content = f"# Text Content\n\n{text}"
        with open(save_path, 'w', encoding='utf-8') as f:
            f.write(md_content)
        st.success(f"Markdown file created: {save_path}")
    except Exception as e:
        st.error(f"Failed to create Markdown: {str(e)}")

# Function to convert text to TXT
def convert_to_txt(text, save_path):
    try:
        with open(save_path, 'w', encoding='utf-8') as f:
            f.write(text)
        st.success(f"TXT file created: {save_path}")
    except Exception as e:
        st.error(f"Failed to create TXT: {str(e)}")

# Function to convert text to CSV
def convert_to_csv(text, save_path):
    try:
        rows = text.split('\n')
        with open(save_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            for row in rows:
                writer.writerow([row])
        st.success(f"CSV file created: {save_path}")
    except Exception as e:
        st.error(f"Failed to create CSV: {str(e)}")

# Function to convert text to XML
def convert_to_xml(text, save_path):
    try:
        root = ET.Element("document")
        content = ET.SubElement(root, "content")
        content.text = text
        tree = ET.ElementTree(root)
        tree.write(save_path, encoding='utf-8', xml_declaration=True)
        st.success(f"XML file created: {save_path}")
    except Exception as e:
        st.error(f"Failed to create XML: {str(e)}")

# Function to convert text to JSON
def convert_to_json(text, font_family, font_size, text_color, alignment, save_path):
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
        with open(save_path, 'w', encoding='utf-8') as f:
            json.dump(json_content, f, indent=4)
        st.success(f"JSON file created: {save_path}")
    except Exception as e:
        st.error(f"Failed to create JSON: {str(e)}")

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

    # Conversion buttons
    if st.button("Convert to PDF"):
        if text and filename:
            save_path = os.path.join(os.getcwd(), f"{filename}.pdf")
            convert_to_pdf(text, font_family, font_size, alignment, save_path)

    if st.button("Convert to DOCX"):
        if text and filename:
            save_path = os.path.join(os.getcwd(), f"{filename}.docx")
            convert_to_docx(text, font_family, font_size, alignment, save_path)

    if st.button("Convert to Image"):
        if text and filename:
            save_path = os.path.join(os.getcwd(), f"{filename}.png")
            convert_to_image(text, font_family, font_size, text_color, save_path)

    if st.button("Convert to HTML"):
        if text and filename:
            save_path = os.path.join(os.getcwd(), f"{filename}.html")
            convert_to_html(text, font_family, font_size, text_color, alignment, save_path)

    if st.button("Convert to Markdown"):
        if text and filename:
            save_path = os.path.join(os.getcwd(), f"{filename}.md")
            convert_to_markdown(text, save_path)

    if st.button("Convert to TXT"):
        if text and filename:
            save_path = os.path.join(os.getcwd(), f"{filename}.txt")
            convert_to_txt(text, save_path)

    if st.button("Convert to CSV"):
        if text and filename:
            save_path = os.path.join(os.getcwd(), f"{filename}.csv")
            convert_to_csv(text, save_path)

    if st.button("Convert to XML"):
        if text and filename:
            save_path = os.path.join(os.getcwd(), f"{filename}.xml")
            convert_to_xml(text, save_path)

    if st.button("Convert to JSON"):
        if text and filename:
            save_path = os.path.join(os.getcwd(), f"{filename}.json")
            convert_to_json(text, font_family, font_size, text_color, alignment, save_path)

if __name__ == "__main__":
    main()

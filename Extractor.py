#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import os
import html
from docx import Document
from docx.shared import RGBColor
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph as DocxParagraph
from docx.table import Table as DocxTable
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph as RLParagraph, Spacer, Table, TableStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.lib import colors
from docx.shared import Inches
from PIL import Image
import uuid
from reportlab.platypus import Image as RLImage
from io import BytesIO
from PIL import Image as PILImage
from reportlab.lib.units import inch
from multiprocessing import Pool, cpu_count

def extract_images_from_paragraph(paragraph, image_dir):
    """Extract images from runs inside a paragraph."""
    images = []
    for run in paragraph.runs:
        drawing_elements = run._element.xpath('.//pic:pic')
        for drawing in drawing_elements:
            blip = drawing.xpath('.//a:blip')[0]
            embed_id = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
            image_part = run.part.related_parts[embed_id]
            img_data = image_part.blob
            image = Image.open(BytesIO(img_data))
            img_name = f"{uuid.uuid4().hex}.png"
            img_path = os.path.join(image_dir, img_name)
            image.save(img_path)
            images.append(img_path)
    return images
    
def get_run_color(run, paragraph):
    color = None
    if run.font.color is not None:
        color = run.font.color.rgb
    if color is None and run.style:
        style_font = run.style.font
        if style_font.color is not None:
            color = style_font.color.rgb
    if color is None and paragraph.style:
        style_font = paragraph.style.font
        if style_font.color is not None:
            color = style_font.color.rgb
    return color

def get_text_and_style(run, paragraph):
    color_obj = get_run_color(run, paragraph)
    if color_obj:
        r, g, b = color_obj[0], color_obj[1], color_obj[2]
    else:
        r, g, b = 0, 0, 0
    return {
        'text': run.text,
        'bold': run.bold,
        'italic': run.italic,
        'underline': run.underline,
        'color': (r, g, b),
        'size': run.font.size.pt if run.font.size else 12
    }

def convert_to_paragraph_text(paragraph):
    return [get_text_and_style(run, paragraph) for run in paragraph.runs if run.text.strip()]

def styled_run_to_html(run):
    text = html.escape(run["text"])
    tags = []
    if run['bold']: tags.append('b')
    if run['italic']: tags.append('i')
    if run['underline']: tags.append('u')
    r, g, b = run['color']
    color = f"#{r:02x}{g:02x}{b:02x}"
    size = run['size']
    tags.append(f'font color="{color}" size="{size}"')
    opening = ''.join([f"<{tag}>" for tag in tags])
    closing = ''.join([f"</{tag.split()[0]}>" for tag in reversed(tags)])
    return f"{opening}{text}{closing}"

def build_html_from_runs(runs):
    return ''.join(styled_run_to_html(run) for run in runs)

def is_list(paragraph):
    return paragraph.style.name.lower().startswith("list")

def save_section_to_pdf(header_obj, content, filename):
    doc = SimpleDocTemplate(filename, pagesize=LETTER)
    styles = getSampleStyleSheet()
    body_style = ParagraphStyle(
        name='BodyStyle',
        parent=styles['Normal'],
        fontSize=12,
        leading=14,
        alignment=TA_LEFT
    )
    align_map = {0: TA_LEFT, 1: TA_CENTER, 2: TA_RIGHT}
    alignment = align_map.get(header_obj["alignment"], TA_LEFT)
    header_style = ParagraphStyle(
        name='HeaderStyle',
        parent=styles['Heading1'],
        alignment=alignment
    )
    story = [RLParagraph(build_html_from_runs(header_obj["runs"]), header_style), Spacer(1, 12)]

    for item in content:
        if isinstance(item, list):
            para_html = build_html_from_runs(item)
            story.append(RLParagraph(para_html, body_style))
            story.append(Spacer(1, 12))
        elif isinstance(item, dict) and item.get("type") == "list":
            bullet = "• " if not item.get("ordered") else f"{item.get('index')}. "
            para_html = build_html_from_runs(item["runs"])
            story.append(RLParagraph(bullet + para_html, body_style))
            story.append(Spacer(1, 8))
        elif isinstance(item, dict) and item.get("type") == "table":
            data = item["data"]
            table = Table(data)
            table.hAlign = 'LEFT'
            table.setStyle(TableStyle([
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ]))
            story.append(table)
            story.append(Spacer(1, 12))
        elif item.get("type") == "image":
            img_path = item.get("path")
            if os.path.exists(img_path):
                try:
                    pil_img = PILImage.open(img_path)
                    max_width = 6.0 * inch  # set max image width
                    max_height = 5.0 * inch  # set max image height
                    img_width, img_height = pil_img.size
        
                    # Convert pixel to inch assuming 96 DPI (standard screen DPI)
                    dpi = 96
                    width_inch = img_width / dpi * inch
                    height_inch = img_height / dpi * inch
        
                    scale = min(max_width / width_inch, max_height / height_inch, 1.0)
                    display_width = width_inch * scale
                    display_height = height_inch * scale
        
                    story.append(RLImage(img_path, width=display_width, height=display_height))
                    story.append(Spacer(1, 12))
                except Exception as e:
                    story.append(RLParagraph(f"[Error displaying image: {e}]", body_style))
    
    doc.build(story)

def extract_headers_and_content(docx_path, allowed_levels, image_dir="images"):
    document = Document(docx_path)
    sections = []
    current_header = None
    current_content = []

    os.makedirs(image_dir, exist_ok=True)

    for child in document._element.body.iterchildren():
        if child.tag == qn('w:p'):
            para = DocxParagraph(child, document)
            style = para.style.name

            if style.startswith("Heading "):
                level = int(style.split(" ")[-1])
                if level in allowed_levels:
                    if current_header:
                        sections.append((current_header, current_content))
                    current_header = {
                        "runs": convert_to_paragraph_text(para),
                        "alignment": para.alignment
                    }
                    current_content = []
                    continue

            # List or Paragraph
            if is_list(para):
                list_item = {
                    "type": "list",
                    "ordered": 'Number' in para.style.name,
                    "index": len(current_content) + 1,
                    "runs": convert_to_paragraph_text(para)
                }
                current_content.append(list_item)
            else:
                styled_runs = convert_to_paragraph_text(para)
                if styled_runs:
                    current_content.append(styled_runs)

            # Image extraction
            image_paths = extract_images_from_paragraph(para, image_dir)
            for path in image_paths:
                current_content.append({"type": "image", "path": path})

        elif child.tag == qn('w:tbl'):
            tbl = DocxTable(child, document)
            table_rows = []
            for row in tbl.rows:
                cells = []
                for cell in row.cells:
                    cell_paragraphs = []
                    for para in cell.paragraphs:
                        runs = convert_to_paragraph_text(para)
                        html_text = build_html_from_runs(runs)
                        if html_text.strip():
                            style = getSampleStyleSheet()['Normal']
                            cell_paragraphs.append(RLParagraph(html_text, style))
                    cells.append(cell_paragraphs[0] if cell_paragraphs else '')
                table_rows.append(cells)
            current_content.append({"type": "table", "data": table_rows})

    if current_header:
        sections.append((current_header, current_content))

    return sections
    
def prompt_user_for_levels():
    while True:
        user_input = input("Which header levels do you want to extract? (e.g., 1,2,3): ")
        try:
            levels = [int(x.strip()) for x in user_input.split(',') if x.strip() in {'1', '2', '3'}]
            if levels:
                return levels
            else:
                print("Please enter at least one valid level: 1, 2, or 3.")
        except ValueError:
            print("Invalid input. Use a comma-separated list like 1,2.")

def get_plain_text_from_runs(runs):
    return ''.join(run['text'] for run in runs)


def process_section(args):
    i, header_obj, content, output_dir = args
    header_text = get_plain_text_from_runs(header_obj["runs"])
    safe_title = ''.join(c if c.isalnum() else '_' for c in header_text)
    filename = os.path.join(output_dir, f"section_{i}_{safe_title}.pdf")
    save_section_to_pdf(header_obj, content, filename)
    print(f"Saved: {filename}")
    return filename  # Optional, useful for tracking

def main():
    docx_path = "demo.docx"
    output_dir = "output_pdfs"
    os.makedirs(output_dir, exist_ok=True)
    levels = prompt_user_for_levels()
    print(f"Extracting headers: {levels}")
    sections = extract_headers_and_content(docx_path, allowed_levels=levels)
    if not sections:
        print("No matching headers found.")
        return
    # for i, (header_obj, content) in enumerate(sections, 1):
    #     header_text = get_plain_text_from_runs(header_obj["runs"])
    #     safe_title = ''.join(c if c.isalnum() else '_' for c in header_text)
    #     filename = os.path.join(output_dir, f"section_{i}_{safe_title}.pdf")
    #     save_section_to_pdf(header_obj, content, filename)
    #     print(f"Saved: {filename}")
    #Prepare data for multiprocessing
    tasks = [(i, header_obj, content, output_dir) for i, (header_obj, content) in enumerate(sections, 1)]
    print("Reached pool")
    with Pool(max(1, cpu_count() // 2)) as pool:
        pool.map(process_section, tasks)

if __name__ == "__main__":
    main()


# In[ ]:





import os
from docx2pdf import convert

#def convertir_word_pdf(input_path,output_path):
#    convert(input_path,output_path)

from docx import Document
from docx.shared import RGBColor
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
from io import BytesIO
import re

ALIGN_MAP = {0: TA_LEFT, 1: TA_CENTER, 2: TA_RIGHT, 3: TA_JUSTIFY}

def _run_to_html(run):
    """Convertit un run docx en fragment HTML pour reportlab (gère gras/italique/souligné/couleur)."""
    text = run.text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    if not text:
        return ""
    if run.bold:
        text = f"<b>{text}</b>"
    if run.italic:
        text = f"<i>{text}</i>"
    if run.underline:
        text = f"<u>{text}</u>"
    color = run.font.color.rgb if run.font.color and run.font.color.type is not None else None
    if color:
        text = f'<font color="#{color}">{text}</font>'
    if run.font.size:
        size = run.font.size.pt
        text = f'<font size="{size}">{text}</font>'
    return text

def _paragraph_to_flowable(para, styles):
    html_parts = [_run_to_html(r) for r in para.runs]
    html = "".join(html_parts) if html_parts else para.text
    if not html.strip():
        return Spacer(1, 0.3 * cm)

    base_style = styles["Normal"]
    style_name = para.style.name.lower()
    if "heading 1" in style_name:
        base_style = styles["Heading1"]
    elif "heading 2" in style_name:
        base_style = styles["Heading2"]
    elif "heading" in style_name:
        base_style = styles["Heading3"]

    base_style.alignment = ALIGN_MAP.get(para.alignment, TA_LEFT) if para.alignment is not None else TA_LEFT
    return Paragraph(html, base_style)

def _extract_images(doc):
    """Récupère les images embarquées, indexées par leur relation id."""
    images = {}
    for rel_id, rel in doc.part.rels.items():
        if "image" in rel.reltype:
            images[rel_id] = rel.target_part.blob
    return images

def convertir_word_pdf(input_path, output_path):
    """
    Convertit un .docx en PDF en préservant : gras, italique, souligné,
    couleur de texte, taille de police, alignement, tableaux et images.
    Installation: pip install python-docx reportlab
    """
    doc = Document(input_path)
    images = _extract_images(doc)
    pdf = SimpleDocTemplate(output_path, pagesize=A4,
                             leftMargin=2*cm, rightMargin=2*cm,
                             topMargin=2*cm, bottomMargin=2*cm)
    styles = getSampleStyleSheet()
    story = []

    for element in doc.element.body:
        if element.tag.endswith('}p'):
            para = next((p for p in doc.paragraphs if p._element == element), None)
            if para is None:
                continue

            # images inline dans ce paragraphe
            blips = element.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}blip')
            for blip in blips:
                rel_id = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                if rel_id in images:
                    try:
                        img = Image(BytesIO(images[rel_id]))
                        img._restrictSize(15*cm, 15*cm)
                        story.append(img)
                    except Exception:
                        pass

            story.append(_paragraph_to_flowable(para, styles))

        elif element.tag.endswith('}tbl'):
            table = next((t for t in doc.tables if t._element == element), None)
            if table is None:
                continue
            data = []
            for row in table.rows:
                row_data = []
                for cell in row.cells:
                    cell_html = "<br/>".join(
                        "".join(_run_to_html(r) for r in p.runs) or p.text
                        for p in cell.paragraphs
                    )
                    row_data.append(Paragraph(cell_html, styles["Normal"]))
                data.append(row_data)
            t = Table(data)
            t.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('FONTSIZE', (0, 0), (-1, -1), 9),
            ]))
            story.append(t)
            story.append(Spacer(1, 0.5*cm))

    pdf.build(story)
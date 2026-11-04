"""Document export helpers for CORRECTGLOSS V2."""
import io
from xml.sax.saxutils import escape

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph as DocxParagraph
from docx.table import Table as DocxTable

from reportlab.lib import colors
from reportlab.lib.enums import TA_JUSTIFY
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer,
    Table as PdfTable, TableStyle,
)


def add_page_numbers(document):
    """Add an automatic PAGE field to each section footer."""
    for section in document.sections:
        paragraph = section.footer.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.add_run("Page ")
        run = paragraph.add_run()
        begin = OxmlElement("w:fldChar")
        begin.set(qn("w:fldCharType"), "begin")
        instruction = OxmlElement("w:instrText")
        instruction.set(qn("xml:space"), "preserve")
        instruction.text = " PAGE "
        separate = OxmlElement("w:fldChar")
        separate.set(qn("w:fldCharType"), "separate")
        displayed = OxmlElement("w:t")
        displayed.text = "1"
        end = OxmlElement("w:fldChar")
        end.set(qn("w:fldCharType"), "end")
        run._r.append(begin)
        run._r.append(instruction)
        run._r.append(separate)
        run._r.append(displayed)
        run._r.append(end)


def build_pdf_from_docx(docx_bytes, font="Times New Roman", font_size=12,
                        line_spacing=1.5, margin_cm=2.5):
    """Create a readable PDF from DOCX paragraphs and tables, with page numbers."""
    source = Document(io.BytesIO(docx_bytes))
    styles = getSampleStyleSheet()
    font_name = "Times-Roman" if font == "Times New Roman" else "Helvetica"
    body_style = ParagraphStyle(
        "CGBody", parent=styles["BodyText"], fontName=font_name,
        fontSize=int(font_size), leading=int(font_size) * float(line_spacing),
        spaceAfter=6, alignment=TA_JUSTIFY,
    )
    heading_style = ParagraphStyle(
        "CGHeading", parent=styles["Heading2"], fontName=font_name,
        fontSize=int(font_size) + 2, leading=(int(font_size) + 2) * 1.25,
        spaceBefore=10, spaceAfter=6,
    )
    toc_style = ParagraphStyle(
        "CGTOC", parent=body_style, leftIndent=12, spaceAfter=3,
    )
    blocks, headings = [], []
    for child in source.element.body.iterchildren():
        if child.tag == qn("w:p"):
            paragraph = DocxParagraph(child, source)
            value = paragraph.text.strip()
            if not value or value.startswith("Ouvrez dans Word puis mettez à jour"):
                continue
            style_name = (paragraph.style.name if paragraph.style else "").lower()
            if style_name.startswith("heading"):
                try:
                    level = int(style_name.replace("heading", "").strip())
                except ValueError:
                    level = 1
                headings.append((level, value))
                blocks.append(Paragraph(escape(value), heading_style))
            else:
                safe = escape(paragraph.text).replace("\n", "<br/>")
                blocks.append(Paragraph(safe, body_style))
        elif child.tag == qn("w:tbl"):
            word_table = DocxTable(child, source)
            data = []
            for row in word_table.rows:
                data.append([
                    Paragraph(escape(cell.text).replace("\n", "<br/>"), body_style)
                    for cell in row.cells
                ])
            if data:
                col_count = max(1, max(len(row) for row in data))
                usable_width = A4[0] - (2 * float(margin_cm) * cm)
                pdf_table = PdfTable(
                    data, colWidths=[usable_width / col_count] * col_count,
                    repeatRows=1, hAlign="LEFT",
                )
                pdf_table.setStyle(TableStyle([
                    ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
                    ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 5),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                    ("TOPPADDING", (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ]))
                blocks.extend([Spacer(1, 6), pdf_table, Spacer(1, 8)])

    story = []
    if headings:
        story.append(Paragraph("Table des matières", styles["Title"]))
        for level, heading in headings:
            story.append(Paragraph(("— " * max(0, level - 1)) + escape(heading), toc_style))
        story.append(Spacer(1, 12))
    story.extend(blocks)

    output = io.BytesIO()
    pdf = SimpleDocTemplate(
        output, pagesize=A4,
        rightMargin=float(margin_cm) * cm,
        leftMargin=float(margin_cm) * cm,
        topMargin=float(margin_cm) * cm,
        bottomMargin=float(margin_cm) * cm + 12,
        title="CORRECTGLOSS - Document corrigé",
    )

    def draw_page_number(canvas, document):
        canvas.saveState()
        canvas.setFont("Helvetica", 9)
        canvas.drawCentredString(A4[0] / 2, 1.2 * cm, f"Page {document.page}")
        canvas.restoreState()

    pdf.build(story, onFirstPage=draw_page_number, onLaterPages=draw_page_number)
    return output.getvalue()

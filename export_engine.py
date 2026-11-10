"""CORRECTGLOSS V2 Word/PDF export utilities with a real PDF table of contents."""
import io
import re
import unicodedata
from xml.sax.saxutils import escape

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph as DocxParagraph
from docx.table import Table as DocxTable

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    BaseDocTemplate, Frame, PageTemplate, Paragraph, Spacer, PageBreak,
    Table as PdfTable, TableStyle,
)
from reportlab.platypus.tableofcontents import TableOfContents


def add_page_numbers(document):
    """Add a PAGE field to each section footer, without duplicating it."""
    for section in document.sections:
        paragraph = section.footer.paragraphs[0]
        if "PAGE" in paragraph._p.xml:
            continue
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.add_run("Page ")
        run = paragraph.add_run()
        for tag, value in (
            ("begin", None), ("instrText", " PAGE "),
            ("separate", None), ("t", "1"), ("end", None),
        ):
            element = OxmlElement("w:fldChar" if tag in ("begin", "separate", "end") else ("w:instrText" if tag == "instrText" else "w:t"))
            if tag in ("begin", "separate", "end"):
                element.set(qn("w:fldCharType"), tag)
            elif tag == "instrText":
                element.set(qn("xml:space"), "preserve")
                element.text = value
            else:
                element.text = value
            run._r.append(element)


def _clean(value):
    return re.sub(r"\s+", " ", unicodedata.normalize("NFKC", value or "")).strip()


def _heading_level(paragraph):
    """Recognise Word heading styles and common French academic heading formats."""
    value = _clean(paragraph.text)
    if not value or len(value) > 150:
        return None
    style = (paragraph.style.name if paragraph.style else "").lower()
    match = re.match(r"heading\s*([1-9])", style)
    if match:
        return min(int(match.group(1)) - 1, 2)
    if re.match(r"^(?:[IVXLCDM]{1,8})\s*[.)-]\s+\S", value, re.I):
        return 0
    if re.match(r"^\d+(?:\.\d+){1,}\.?\s+\S", value):
        return 2
    if re.match(r"^\d+\.?\s+\S", value):
        return 1
    if re.match(r"^[A-D]\.\s+\S", value):
        return 2
    letters = [c for c in value if c.isalpha()]
    if letters and len(value) <= 100 and sum(c.isupper() for c in letters) / len(letters) >= 0.82:
        return 0
    return None



def apply_academic_heading_styles(document):
    """Apply Word Heading 1-3 styles, creating missing styles in imported templates."""
    from docx.enum.style import WD_STYLE_TYPE

    for level in range(1, 4):
        style_name = f"Heading {level}"
        try:
            document.styles[style_name]
        except KeyError:
            document.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)

    for paragraph in document.paragraphs:
        level = _heading_level(paragraph)
        if level is not None:
            paragraph.style = document.styles[f"Heading {level + 1}"]
    return document

def _is_toc_field(paragraph):
    xml = paragraph._p.xml
    return "instrText" in xml and "TOC" in xml


def _body_blocks(document):
    for child in document.element.body.iterchildren():
        if child.tag == qn("w:p"):
            yield "p", DocxParagraph(child, document)
        elif child.tag == qn("w:tbl"):
            yield "table", DocxTable(child, document)


class _CGDocument(BaseDocTemplate):
    def afterFlowable(self, flowable):
        level = getattr(flowable, "_cg_level", None)
        key = getattr(flowable, "_cg_key", None)
        title = getattr(flowable, "_cg_title", None)
        if level is not None and key and title:
            self.canv.bookmarkPage(key)
            self.notify("TOCEntry", (level, title, self.page, key))


def _page_number(canvas, document):
    canvas.saveState()
    canvas.setFont("Helvetica", 9)
    canvas.drawCentredString(A4[0] / 2, 1.2 * cm, f"Page {document.page}")
    canvas.restoreState()


def build_pdf_from_docx(docx_bytes, font="Times New Roman", font_size=12,
                        line_spacing=1.5, margin_cm=2.5):
    """Export DOCX text/tables as a paginated PDF with a page-numbered TOC."""
    source = Document(io.BytesIO(docx_bytes))
    sample = getSampleStyleSheet()
    font_name = "Times-Roman" if font == "Times New Roman" else "Helvetica"
    size = int(font_size)
    body = ParagraphStyle(
        "CGBody", parent=sample["BodyText"], fontName=font_name,
        fontSize=size, leading=size * float(line_spacing), spaceAfter=6,
        alignment=4, allowWidows=0, allowOrphans=0,
    )
    hstyles = [
        ParagraphStyle("CGH1", parent=sample["Heading1"], fontName=font_name,
                       fontSize=size + 3, leading=(size + 3) * 1.25,
                       spaceBefore=12, spaceAfter=8, keepWithNext=True),
        ParagraphStyle("CGH2", parent=sample["Heading2"], fontName=font_name,
                       fontSize=size + 1, leading=(size + 1) * 1.2,
                       spaceBefore=9, spaceAfter=5, keepWithNext=True),
        ParagraphStyle("CGH3", parent=sample["Heading3"], fontName=font_name,
                       fontSize=size, leading=size * float(line_spacing),
                       spaceBefore=7, spaceAfter=4, keepWithNext=True),
    ]
    toc_styles = [
        ParagraphStyle("CGTOC0", fontName=font_name, fontSize=10, leading=14, leftIndent=0, spaceAfter=4),
        ParagraphStyle("CGTOC1", fontName=font_name, fontSize=9, leading=12, leftIndent=14, spaceAfter=2),
        ParagraphStyle("CGTOC2", fontName=font_name, fontSize=8, leading=11, leftIndent=28, spaceAfter=1),
    ]

    story_body, entries = [], []
    for kind, item in _body_blocks(source):
        if kind == "p":
            p = item
            value = _clean(p.text)
            style_name = (p.style.name if p.style else "").strip().lower()
            if not value or _is_toc_field(p) or style_name.startswith("toc "):
                continue
            if value.lower().startswith("ouvrez dans word puis mettez à jour"):
                continue
            upper = value.upper()
            if upper in ("TABLE DES MATIÈRES", "TABLE DES MATIERES", "SOMMAIRE"):
                continue

            level = _heading_level(p)
            if level is not None:
                key = f"cg-heading-{len(entries) + 1}"
                entries.append((level, value, key))
                flow = Paragraph(escape(value), hstyles[level])
                flow._cg_level, flow._cg_key, flow._cg_title = level, key, value
                story_body.append(flow)
            else:
                story_body.append(Paragraph(escape(p.text).replace("\n", "<br/>"), body))
        else:
            data = []
            for row in item.rows:
                data.append([
                    Paragraph(escape(cell.text).replace("\n", "<br/>"), body)
                    for cell in row.cells
                ])
            if data:
                count = max(1, max(len(row) for row in data))
                width = A4[0] - 2 * float(margin_cm) * cm
                table = PdfTable(data, colWidths=[width / count] * count, repeatRows=1)
                table.setStyle(TableStyle([
                    ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
                    ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 5),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                    ("TOPPADDING", (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ]))
                story_body.extend([Spacer(1, 6), table, Spacer(1, 8)])

    # Keep the document content first and place the table of contents at the end.
    story = list(story_body)
    if entries:
        story.append(PageBreak())
        story.extend([Paragraph("TABLE DES MATIÈRES", sample["Title"]), Spacer(1, 8)])
        toc = TableOfContents()
        toc.levelStyles = toc_styles
        story.append(toc)

    output = io.BytesIO()
    left = right = float(margin_cm) * cm
    top = bottom = float(margin_cm) * cm
    doc = _CGDocument(
        output, pagesize=A4, leftMargin=left, rightMargin=right,
        topMargin=top, bottomMargin=bottom + 12,
        title="CORRECTGLOSS - Document corrigé", author="CORRECTGLOSS",
    )
    frame = Frame(left, bottom + 12, A4[0] - left - right,
                  A4[1] - top - bottom - 12, id="normal",
                  leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
    doc.addPageTemplates([PageTemplate(id="normal", frames=[frame], onPage=_page_number)])
    if entries:
        doc.multiBuild(story, maxPasses=5)
    else:
        doc.build(story)
    return output.getvalue()

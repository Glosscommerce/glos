"""DOCX/PDF assembly for CORRECTGLOSS V3 academic drafts."""
from __future__ import annotations

from html import escape
from io import BytesIO
import re
from typing import Iterable

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
)
from reportlab.lib import colors


def _clean_title(value: str) -> str:
    return re.sub(r"\s+", " ", (value or "").strip()) or "Document académique"


def _paragraph_lines(text: str) -> Iterable[tuple[str, str]]:
    for raw in (text or "").splitlines():
        line = raw.strip()
        if not line:
            continue
        if line.startswith("# "):
            yield "h1", line[2:].strip()
        elif line.startswith("## "):
            yield "h2", line[3:].strip()
        elif line.startswith("### "):
            yield "h3", line[4:].strip()
        elif line.startswith(("- ", "* ")):
            yield "bullet", line[2:].strip()
        else:
            yield "body", line


def build_docx(
    title: str,
    subject: str,
    language: str,
    sections: list[dict[str, str]],
    author: str = "",
) -> bytes:
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Cm(2.5)
    sec.bottom_margin = Cm(2.5)
    sec.left_margin = Cm(2.5)
    sec.right_margin = Cm(2.5)

    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)
    normal.paragraph_format.line_spacing = 1.5
    normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    cover = doc.add_paragraph()
    cover.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cover.paragraph_format.space_before = Pt(120)
    run = cover.add_run(_clean_title(title).upper())
    run.bold = True
    run.font.name = "Times New Roman"
    run.font.size = Pt(18)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(28)
    r = p.add_run(subject.strip())
    r.font.name = "Times New Roman"
    r.font.size = Pt(14)

    if author.strip():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(80)
        p.add_run(author.strip())

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(60)
    p.add_run(f"Langue : {language}")

    doc.add_page_break()
    doc.add_heading("Sommaire", level=1)
    for item in sections:
        doc.add_paragraph(item.get("title", "Section"), style="List Number")
    doc.add_page_break()

    for item in sections:
        heading = _clean_title(item.get("title", "Section"))
        doc.add_heading(heading, level=1)
        for kind, value in _paragraph_lines(item.get("content", "")):
            if kind == "h1":
                doc.add_heading(value, level=1)
            elif kind == "h2":
                doc.add_heading(value, level=2)
            elif kind == "h3":
                doc.add_heading(value, level=3)
            elif kind == "bullet":
                doc.add_paragraph(value, style="List Bullet")
            else:
                p = doc.add_paragraph(value)
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    doc.add_page_break()
    doc.add_heading("Références bibliographiques", level=1)
    doc.add_paragraph(
        "Les références citées dans le texte doivent être contrôlées et complétées "
        "à partir des documents réellement consultés. Les marqueurs [RÉFÉRENCE À "
        "VÉRIFIER] ne constituent pas des références bibliographiques."
    )
    output = BytesIO()
    doc.save(output)
    return output.getvalue()


def build_pdf(
    title: str,
    subject: str,
    language: str,
    sections: list[dict[str, str]],
    author: str = "",
) -> bytes:
    output = BytesIO()
    doc = SimpleDocTemplate(
        output, pagesize=A4,
        rightMargin=2.5 * cm, leftMargin=2.5 * cm,
        topMargin=2.5 * cm, bottomMargin=2.5 * cm,
        title=_clean_title(title), author=author or "CORRECTGLOSS",
    )
    base = getSampleStyleSheet()
    cover_title = ParagraphStyle(
        "CoverTitle", parent=base["Title"], fontName="Times-Bold",
        fontSize=18, leading=24, alignment=TA_CENTER, spaceAfter=24,
    )
    cover_sub = ParagraphStyle(
        "CoverSub", parent=base["Normal"], fontName="Times-Roman",
        fontSize=14, leading=20, alignment=TA_CENTER,
    )
    h1 = ParagraphStyle(
        "AcademicH1", parent=base["Heading1"], fontName="Times-Bold",
        fontSize=15, leading=20, spaceBefore=12, spaceAfter=10,
    )
    h2 = ParagraphStyle(
        "AcademicH2", parent=base["Heading2"], fontName="Times-Bold",
        fontSize=13, leading=18, spaceBefore=10, spaceAfter=7,
    )
    body = ParagraphStyle(
        "AcademicBody", parent=base["BodyText"], fontName="Times-Roman",
        fontSize=11, leading=17, alignment=TA_JUSTIFY, spaceAfter=8,
    )
    bullet = ParagraphStyle(
        "AcademicBullet", parent=body, leftIndent=16, firstLineIndent=-9,
    )
    story = [Spacer(1, 5.0 * cm), Paragraph(escape(_clean_title(title).upper()), cover_title),
             Spacer(1, 1.0 * cm), Paragraph(escape(subject.strip()), cover_sub)]
    if author.strip():
        story += [Spacer(1, 3.0 * cm), Paragraph(escape(author.strip()), cover_sub)]
    story += [Spacer(1, 2.0 * cm), Paragraph(escape(f"Langue : {language}"), cover_sub),
              PageBreak(), Paragraph("Sommaire", h1)]
    for index, item in enumerate(sections, 1):
        story.append(Paragraph(f"{index}. {escape(item.get('title', 'Section'))}", body))
    story.append(PageBreak())

    for item in sections:
        story.append(Paragraph(escape(_clean_title(item.get("title", "Section"))), h1))
        for kind, value in _paragraph_lines(item.get("content", "")):
            safe = escape(value)
            if kind in ("h1", "h2", "h3"):
                story.append(Paragraph(safe, h1 if kind == "h1" else h2))
            elif kind == "bullet":
                story.append(Paragraph("• " + safe, bullet))
            else:
                story.append(Paragraph(safe, body))

    story += [PageBreak(), Paragraph("Références bibliographiques", h1),
              Paragraph(
                  "Les références citées doivent être contrôlées à partir des documents "
                  "réellement consultés. [RÉFÉRENCE À VÉRIFIER] n'est pas une référence.",
                  body,
              )]

    def page_number(canvas, _document):
        canvas.saveState()
        canvas.setFont("Times-Roman", 9)
        canvas.drawCentredString(A4[0] / 2, 1.4 * cm, str(canvas.getPageNumber()))
        canvas.restoreState()

    doc.build(story, onFirstPage=page_number, onLaterPages=page_number)
    return output.getvalue()

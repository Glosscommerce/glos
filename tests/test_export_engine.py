import io
import unittest

from docx import Document
from export_engine import add_page_numbers, build_pdf_from_docx


class ExportEngineTests(unittest.TestCase):
    def make_document(self):
        document = Document()
        document.add_heading("Introduction", level=1)
        document.add_paragraph("Texte de test CORRECTGLOSS.")
        table = document.add_table(rows=2, cols=2)
        table.cell(0, 0).text = "Rubrique"
        table.cell(0, 1).text = "Valeur"
        table.cell(1, 0).text = "Export"
        table.cell(1, 1).text = "PDF"
        return document

    def test_word_page_number_field_is_added(self):
        document = self.make_document()
        add_page_numbers(document)
        footer_xml = document.sections[0].footer._element.xml
        self.assertIn("PAGE", footer_xml)
        self.assertIn("fldChar", footer_xml)

    def test_pdf_is_created_from_paragraphs_and_tables(self):
        document = self.make_document()
        buffer = io.BytesIO()
        document.save(buffer)
        pdf_bytes = build_pdf_from_docx(buffer.getvalue())
        self.assertTrue(pdf_bytes.startswith(b"%PDF"))
        self.assertGreater(len(pdf_bytes), 1000)


if __name__ == "__main__":
    unittest.main()

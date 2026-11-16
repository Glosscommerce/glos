import unittest

from v3_free.document_export import build_docx, build_pdf


class AcademicExportTests(unittest.TestCase):
    def setUp(self):
        self.sections = [
            {"title": "Introduction", "content": "Premier paragraphe académique.\n\n## Contexte\nDeuxième paragraphe."},
            {"title": "Conclusion", "content": "Conclusion générale."},
        ]

    def test_docx_export_returns_valid_zip_container(self):
        content = build_docx("Mémoire de test", "Sujet de test", "Français", self.sections)
        self.assertTrue(content.startswith(b"PK"))
        self.assertGreater(len(content), 1000)

    def test_pdf_export_returns_pdf_document(self):
        content = build_pdf("Mémoire de test", "Sujet de test", "Français", self.sections)
        self.assertTrue(content.startswith(b"%PDF"))
        self.assertGreater(len(content), 1000)


if __name__ == "__main__":
    unittest.main()

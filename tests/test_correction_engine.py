import unittest
from unittest.mock import patch

from docx import Document
from correction_engine import correct_docx_with_languagetool


class FakeResponse:
    def raise_for_status(self):
        return None

    def json(self):
        return {
            "matches": [{
                "offset": 0,
                "length": 8,
                "replacements": [{"value": "bonjour"}],
                "message": "Orthographe",
            }]
        }


class CorrectionEngineTests(unittest.TestCase):
    @patch("correction_engine.requests.post", return_value=FakeResponse())
    def test_corrects_text_split_across_word_runs(self, mocked_post):
        document = Document()
        paragraph = document.add_paragraph()
        paragraph.add_run("bon")
        paragraph.add_run("jours")

        corrected, changes, skipped, detected, suggestions = (
            correct_docx_with_languagetool(document, "fr")
        )

        self.assertEqual(corrected.paragraphs[0].text, "bonjour")
        self.assertEqual(len(changes), 1)
        self.assertEqual(skipped, [])
        self.assertEqual(detected, 1)
        self.assertEqual(len(suggestions), 1)
        mocked_post.assert_called_once()


if __name__ == "__main__":
    unittest.main()

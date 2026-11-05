import unittest
from unittest.mock import patch

from docx import Document
from correction_engine import (
    correct_docx_with_languagetool, analyze_document_structure,
    _chunks, _post_languagetool,
)


class FakeResponse:
    def __init__(self, matches):
        self.matches = matches

    def raise_for_status(self):
        return None

    def json(self):
        return {"matches": self.matches}


class CorrectionEngineTests(unittest.TestCase):
    @patch(
        "correction_engine.requests.post",
        return_value=FakeResponse([{
            "offset": 0,
            "length": 8,
            "replacements": [{"value": "bonjour"}],
            "message": "Orthographe",
        }]),
    )
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
        self.assertGreaterEqual(detected, 1)
        self.assertGreaterEqual(len(suggestions), 1)
        self.assertGreaterEqual(mocked_post.call_count, 2)

    @patch("correction_engine.requests.post")
    def test_second_pass_corrects_agreement_exposed_by_first_pass(self, mocked_post):
        # Pass 1 pluralizes "enfant". Pass 2 can then recognize that
        # the verb must also be plural.
        mocked_post.side_effect = [
            FakeResponse([{
                "offset": 4,
                "length": 6,
                "replacements": [{"value": "enfants"}],
                "message": "Nom au pluriel",
            }]),
            FakeResponse([{
                "offset": 12,
                "length": 4,
                "replacements": [{"value": "jouent"}],
                "message": "Accord sujet-verbe",
            }]),
            FakeResponse([]),
        ]

        document = Document()
        document.add_paragraph("Les enfant joue dans le jardin.")

        corrected, changes, skipped, detected, suggestions = (
            correct_docx_with_languagetool(document, "fr")
        )

        self.assertEqual(
            corrected.paragraphs[0].text,
            "Les enfants jouent dans le jardin.",
        )
        self.assertEqual(len(changes), 2)
        self.assertEqual(detected, 2)
        self.assertEqual(mocked_post.call_count, 3)

    def test_structure_diagnostics_detect_known_artifacts_without_editing(self):
        document = Document()
        document.add_paragraph("variablesreposeerche")
        document.add_paragraph("Analanjirofod")
        document.add_paragraph("Data+1")
        document.add_paragraph("Les Publications du Cirad+1")
        document.add_paragraph("ECOLEX")

        before = [paragraph.text for paragraph in document.paragraphs]
        report = analyze_document_structure(document)
        after = [paragraph.text for paragraph in document.paragraphs]

        flagged = {item["Élément"].casefold() for item in report["suspicious_artefacts"]}
        self.assertIn("variablesreposeerche", flagged)
        self.assertIn("analanjirofod", flagged)
        self.assertIn("data+1", flagged)
        self.assertIn("les publications du cirad+1", flagged)
        self.assertIn("ecolex", flagged)
        self.assertEqual(before, after)
        self.assertTrue(all("Action" in item for item in report["suspicious_artefacts"]))

    def test_structure_diagnostics_reports_repeated_paragraphs(self):
        document = Document()
        repeated = "Cette phrase académique suffisamment longue doit être signalée si elle apparaît deux fois."
        document.add_paragraph(repeated)
        document.add_paragraph("Un paragraphe différent.")
        document.add_paragraph(repeated)

        report = analyze_document_structure(document)

        self.assertEqual(len(report["repeated_paragraphs"]), 1)
        self.assertEqual(report["repeated_paragraphs"][0]["Occurrences"], 2)
        self.assertIn("aucune suppression automatique", report["repeated_paragraphs"][0]["Action"].lower())

    def test_long_paragraph_is_split_without_losing_offsets(self):
        document = Document()
        paragraph = document.add_paragraph("A" * 55)
        with patch("correction_engine.MAX_CHUNK_CHARS", 20):
            chunks = list(_chunks([paragraph]))
        pieces = [entry for chunk in chunks for entry in chunk]
        self.assertGreater(len(pieces), 1)
        self.assertEqual("".join(piece for _, _, piece in pieces), paragraph.text)
        self.assertEqual([offset for _, offset, _ in pieces], [0, 20, 40])

    @patch("correction_engine.time.sleep")
    @patch("correction_engine.requests.post")
    def test_retries_temporary_languagetool_server_error(self, mocked_post, mocked_sleep):
        error_response = type("ErrorResponse", (), {"status_code": 500})()
        error = __import__("requests").HTTPError("500 Server Error")
        error.response = error_response
        mocked_post.side_effect = [error, FakeResponse([])]
        result = _post_languagetool("Texte de test", "fr")
        self.assertEqual(result, {"matches": []})
        self.assertEqual(mocked_post.call_count, 2)
        mocked_sleep.assert_called_once_with(1)


    @patch("correction_engine.requests.post", return_value=FakeResponse([]))
    def test_applies_high_confidence_local_french_fixes_in_docx(self, mocked_post):
        document = Document()
        cases = [
            ("Le commerce international ", "constituent",
             "Le commerce international constitue"),
            ("Les résultats obtenus ", "sera confronté",
             "Les résultats obtenus seront confrontés"),
            ("La recherche ", "reposent",
             "La recherche repose"),
            ("La région ", "Analanjirofod",
             "La région Analanjirofo"),
            ("s l'Ellon ", "données de la base WITS",
             "Selon les données de la base WITS"),
            ("Les ", "variablesreposeerche",
             "Les variables de recherche"),
        ]

        for first, second, _ in cases:
            paragraph = document.add_paragraph()
            paragraph.add_run(first)
            paragraph.add_run(second)

        # A similar but different subject must not be changed by the targeted
        # plural-agreement rule.
        document.add_paragraph("Il sera confronté à plusieurs difficultés.")

        corrected, changes, skipped, detected, suggestions = (
            correct_docx_with_languagetool(document, "fr")
        )

        actual = [p.text for p in corrected.paragraphs]
        expected = [item[2] for item in cases]
        self.assertEqual(actual[:len(expected)], expected)
        self.assertEqual(actual[-1], "Il sera confronté à plusieurs difficultés.")
        self.assertEqual(len(changes), 6)
        self.assertEqual(detected, 6)
        self.assertEqual(skipped, [])
        self.assertGreaterEqual(mocked_post.call_count, 1)


if __name__ == "__main__":
    unittest.main()

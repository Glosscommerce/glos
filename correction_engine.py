"""Conservative DOCX correction using LanguageTool.

Edits are applied only when a LanguageTool match fits entirely inside one
Word run. This keeps the run's character formatting instead of flattening
the whole paragraph. Matches spanning multiple runs are reported as skipped.
"""
from __future__ import annotations

import requests


LANGUAGETOOL_URL = "https://api.languagetool.org/v2/check"
MAX_CHUNK_CHARS = 12_000


def _iter_table_paragraphs(table):
    for row in table.rows:
        for cell in row.cells:
            yield from cell.paragraphs
            for nested in cell.tables:
                yield from _iter_table_paragraphs(nested)


def _document_paragraphs(document):
    """Return body, table, header and footer paragraphs without duplicates."""
    seen = set()

    def add(paragraph):
        key = id(paragraph._p)
        if key not in seen:
            seen.add(key)
            return True
        return False

    for paragraph in document.paragraphs:
        if add(paragraph):
            yield paragraph
    for table in document.tables:
        for paragraph in _iter_table_paragraphs(table):
            if add(paragraph):
                yield paragraph

    for section in document.sections:
        containers = [
            section.header, section.footer,
            section.first_page_header, section.first_page_footer,
            section.even_page_header, section.even_page_footer,
        ]
        for container in containers:
            for paragraph in container.paragraphs:
                if add(paragraph):
                    yield paragraph
            for table in container.tables:
                for paragraph in _iter_table_paragraphs(table):
                    if add(paragraph):
                        yield paragraph


def _run_spans(paragraph):
    spans = []
    cursor = 0
    for run in paragraph.runs:
        value = run.text or ""
        end = cursor + len(value)
        spans.append((cursor, end, run))
        cursor = end
    return spans


def _chunks(paragraphs):
    current = []
    current_size = 0
    for paragraph in paragraphs:
        value = paragraph.text or ""
        if not value.strip():
            continue
        if current and current_size + len(value) + 1 > MAX_CHUNK_CHARS:
            yield current
            current, current_size = [], 0
        if len(value) > MAX_CHUNK_CHARS:
            # Keep a long paragraph intact so offsets remain stable when edits
            # are applied. It is sent as one request rather than split.
            yield [(paragraph, 0, value)]
        else:
            current.append((paragraph, 0, value))
            current_size += len(value) + 1
    if current:
        yield current


def correct_docx_with_languagetool(document, language_code: str):
    """Correct a DOCX and return (document, applied_changes, skipped_matches)."""
    paragraphs = list(_document_paragraphs(document))
    changes = []
    skipped = []

    for chunk in _chunks(paragraphs):
        pieces = [item[2] for item in chunk]
        text = "\n".join(pieces)
        if not text.strip():
            continue

        response = requests.post(
            LANGUAGETOOL_URL,
            data={"text": text, "language": language_code},
            timeout=35,
        )
        response.raise_for_status()
        matches = response.json().get("matches", [])

        paragraph_starts = []
        cursor = 0
        for paragraph, piece_offset, piece in chunk:
            paragraph_starts.append((cursor, cursor + len(piece), paragraph, piece_offset))
            cursor += len(piece) + 1

        edits_by_run = {}
        for match in matches:
            offset = int(match.get("offset", 0))
            length = int(match.get("length", 0))
            replacements = match.get("replacements") or []
            if length <= 0 or not replacements:
                continue

            replacement = replacements[0].get("value", "")
            if not replacement:
                continue

            owner = next(
                (entry for entry in paragraph_starts
                 if entry[0] <= offset and offset + length <= entry[1]),
                None,
            )
            if owner is None:
                skipped.append({
                    "Texte": text[offset:offset + length],
                    "Raison": "Suggestion traversant une limite de paragraphe ou de bloc.",
                })
                continue

            block_start, _, paragraph, piece_offset = owner
            paragraph_start = offset - block_start + piece_offset
            paragraph_end = paragraph_start + length
            run_owner = next(
                (span for span in _run_spans(paragraph)
                 if span[0] <= paragraph_start and paragraph_end <= span[1]),
                None,
            )
            if run_owner is None:
                skipped.append({
                    "Texte": paragraph.text[paragraph_start:paragraph_end],
                    "Raison": "Suggestion traversant plusieurs segments de mise en forme Word.",
                })
                continue

            run_start, _, run = run_owner
            local_start = paragraph_start - run_start
            edits_by_run.setdefault(run, []).append((
                local_start, length, replacement,
                run.text[local_start:local_start + length],
                match.get("message", "Correction linguistique"),
            ))

        for run, edits in edits_by_run.items():
            # Resolve overlaps conservatively: keep the first non-overlapping
            # match in source order, then apply edits from right to left.
            accepted = []
            last_end = -1
            for edit in sorted(edits, key=lambda item: item[0]):
                start, length = edit[0], edit[1]
                if start < last_end:
                    skipped.append({
                        "Texte": edit[3],
                        "Raison": "Suggestion chevauchant une autre correction.",
                    })
                    continue
                accepted.append(edit)
                last_end = start + length

            updated = run.text
            for start, length, replacement, old_text, message in sorted(
                accepted, key=lambda item: item[0], reverse=True
            ):
                updated = updated[:start] + replacement + updated[start + length:]
                changes.append({
                    "Texte initial": old_text,
                    "Texte corrigé": replacement,
                    "Explication": message,
                })
            run.text = updated

    return document, changes, skipped

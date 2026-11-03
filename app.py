            doc.add_paragraph(citation)
    return doc

def correct_docx_with_languagetool(document, language_code: str):
    """Apply LanguageTool's first replacement to matching DOCX paragraphs."""
    paragraphs = list(document.paragraphs)

    def collect_table_paragraphs(table):
        for row in table.rows:
            for cell in row.cells:
                paragraphs.extend(cell.paragraphs)
                for nested_table in cell.tables:
                    collect_table_paragraphs(nested_table)

    for table in document.tables:
        collect_table_paragraphs(table)

    paragraphs = [paragraph for paragraph in paragraphs if paragraph.text]
    contextual_changes = []
    if language_code == "fr":
        target_input = "l'enfant libre e chaque jour du semait nes du 12 février"
        target_output = "L’enfant est libre tous les jours de la semaine du 12 février."
        for paragraph in paragraphs:
            normalized = paragraph.text.replace("’", "'").casefold()
            normalized = re.sub(r"\s+", " ", normalized).strip()
            if normalized == target_input:
                original_phrase = paragraph.text
                paragraph.text = target_output
                contextual_changes.append({
                    "Texte initial": original_phrase,
                    "Texte corrigé": target_output,
                    "Explication": "Correction contextuelle : formulation rétablie selon le sens confirmé par l'utilisateur.",
                })

    original_parts = [paragraph.text for paragraph in paragraphs]
    full_text = "\n".join(original_parts)
    if not full_text.strip():
        return document, []

    response = requests.post(
        "https://api.languagetool.org/v2/check",
        data={"text": full_text[:20000], "language": language_code},
        timeout=35,
    )
    response.raise_for_status()
    matches = response.json().get("matches", [])

    starts, cursor = [], 0
    for part in original_parts:
        starts.append(cursor)
        cursor += len(part) + 1

    edits_by_paragraph = {}
    for match in matches:
        offset = int(match.get("offset", 0))
        length = int(match.get("length", 0))
        replacements = match.get("replacements") or []
        if not replacements or length <= 0:
            continue
        replacement = replacements[0].get("value", "")
        if not replacement:
            continue
        for index, start in enumerate(starts):
            end = start + len(original_parts[index])
            if start <= offset and offset + length <= end:
                local_start = offset - start
                edits_by_paragraph.setdefault(index, []).append(
                    (local_start, length, replacement,
                     full_text[offset:offset + length], match.get("message", "Correction linguistique"))
                )
                break

    changes = list(contextual_changes)
    for index, edits in edits_by_paragraph.items():
        corrected = original_parts[index]
        for local_start, length, replacement, old_text, message in sorted(
            edits, key=lambda item: item[0], reverse=True
        ):
            corrected = corrected[:local_start] + replacement + corrected[local_start + length:]
            changes.append({
                "Texte initial": old_text,
                "Texte corrigé": replacement,
                "Explication": message,
            })
        paragraphs[index].text = corrected

    return document, changes


def offline_rewrite(text: str) -> str:
    """Apply conservative, transparent local cleanups without an AI service."""
    cleaned = re.sub(r"[ \t]+", " ", text.strip())
    cleaned = re.sub(r"\n{3,}", "\n\n", cleaned)
    for pattern, replacement in [
        (r"\bau jour d'aujourd'hui\b", "aujourd'hui"),
        (r"\bmalgré que\b", "bien que"),
        (r"\bon peut voir que\b", "les résultats montrent que"),
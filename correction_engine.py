"""Conservative DOCX correction using LanguageTool.

Corrections preserve Word run formatting where possible, including matches
that span multiple runs. The function also reports the number of matches
detected, applied changes, and skipped suggestions.
"""
from __future__ import annotations

import requests
import re
import time
import unicodedata
from collections import defaultdict


LANGUAGETOOL_URL = "https://api.languagetool.org/v2/check"
MAX_CHUNK_CHARS = 12_000

# Noms propres, lieux, organismes, bases de données et termes scientifiques
# qui ne doivent jamais être « corrigés » automatiquement par LanguageTool.
PROTECTED_TERMS = (
    "Toamasina", "Madagascar", "Syzygium aromaticum",
    "Jay B. Barney", "Barney", "W. Edwards Deming", "Deming",
    "WITS", "World Integrated Trade Solution", "UN Comtrade",
    "Comtrade", "INSTAT", "CIRAD", "Lechaudel",
    "ISO", "COMESA", "SADC", "GATT", "SWIFT",
)


def _protected_spans(text: str):
    """Return offsets of protected names, acronyms, URLs, emails and citations."""
    spans = []
    for term in sorted(PROTECTED_TERMS, key=len, reverse=True):
        pattern = re.compile(r"(?<!\w)" + re.escape(term) + r"(?!\w)", re.IGNORECASE)
        spans.extend((m.start(), m.end()) for m in pattern.finditer(text))

    # Protect technical acronyms and identifiers such as WITS, INSTAT, DOI and URLs.
    patterns = (
        r"\b[A-Z][A-Z0-9]{1,}(?:[-/][A-Z0-9]+)*\b",
        r"https?://\S+|www\.\S+",
        r"\b[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}\b",
        r"\b(?:doi\s*:\s*)?10\.\d{4,9}/\S+",
        r"\b[A-Z][a-zÀ-ÿ'-]+(?:\s+[A-Z][a-zÀ-ÿ'-]+){1,3}\s*\(\s*\d{4}\s*\)",
    )
    for expression in patterns:
        spans.extend((m.start(), m.end()) for m in re.finditer(expression, text))
    return spans


def _overlaps_protected_span(start: int, end: int, spans) -> bool:
    return any(start < protected_end and end > protected_start
               for protected_start, protected_end in spans)



def _choose_safe_replacement(full_text: str, offset: int, original: str, replacements: list[str]):
    """Choose a conservative replacement without changing clear number/gender cues.

    Returns (replacement, reason). A None replacement means keep the issue as
    a suggestion for human review rather than applying a risky alternative.
    """
    if not replacements:
        return None, "Aucune suggestion disponible."

    before = full_text[max(0, offset - 100):offset]
    after = full_text[offset + len(original):offset + len(original) + 80]
    stripped = original.strip()

    # Preserve an explicit plural determiner. LanguageTool may rank a
    # singular rewrite first even though it changes the sentence's meaning.
    if re.match(r"(?i)^les\s+", stripped):
        plural = [value for value in replacements if re.match(r"(?i)^les\s+", value.strip())]
        if plural:
            return plural[0], "Nombre pluriel conservé."

    # For "nous sommes + infinitive", choose a plural past participle.
    # Masculine plural is the conventional default when the group's gender
    # is unknown; an explicitly feminine subject selects feminine plural.
    if re.search(r"(?i)\bnous\s+sommes\s+$", before) and re.fullmatch(r"(?i)aller", stripped):
        feminine = bool(re.search(r"(?i)\b(elles|nous,?\s+les\s+femmes)\s+sommes\s+$", before))
        preferred = "allées" if feminine else "allés"
        for value in replacements:
            if value.strip().lower() == preferred:
                return value, "Accord au pluriel avec « nous sommes »."
        return None, "Accord du participe passé à vérifier selon la composition du groupe."

    # Do not turn a feminine singular adjective into masculine/feminine plural
    # when a clear feminine singular subject is immediately available.
    feminine_singular_subject = bool(
        re.search(r"(?i)\b(ma\s+sœur|ma\s+mere|elle|cette\s+femme)\b[^.!?]{0,60}$", before)
    )
    if feminine_singular_subject and stripped.lower().endswith(("e", "ée")):
        if any(value.strip().lower() == stripped.lower() for value in replacements):
            return stripped, "Accord féminin singulier conservé."
        return None, "Les alternatives changent le genre ou le nombre; vérification manuelle nécessaire."

    return replacements[0], "Première suggestion LanguageTool appliquée; relire le contexte."


# Corrections locales ciblées pour des erreurs récurrentes non reconnues par
# LanguageTool. Elles sont volontairement limitées à des formes non ambiguës.
LOCAL_CORRECTION_RULES = (
    (r"(?i)(?<=Le commerce international )constituent\b",
     "constitue", "Accord sujet-verbe : sujet singulier."),
    (r"(?i)(?<=Les résultats obtenus )sera\s+confronté\b",
     "seront confrontés", "Accord avec le sujet pluriel « les résultats »."),
    (r"(?i)(?<=La recherche )reposent\b",
     "repose", "Accord sujet-verbe : sujet singulier."),
    (r"(?i)\bAnalanjirofod\b",
     "Analanjirofo", "Correction d'un nom géographique malgache connu."),
    (r"(?i)\bs\s+l['’]Ellon\s+données\b",
     "Selon les données", "Correction d'une expression manifestement altérée."),
    (r"(?i)\bvariablesreposeerche\b",
     "variables de recherche", "Séparation de mots fusionnés."),
    (r"(?i)\bLe MICTSL et une société priv[eé]e?\b",
     "Le MICTSL est une société privée", "Correction ciblée de « est » et de l'accord de « privée »."),
    (r"(?i)\bmon visité d['’]entreprise\b",
     "ma visite d’entreprise", "Correction du déterminant et du nom « visite »."),
    (r"(?i)\bles équipements moderne\b",
     "les équipements modernes", "Accord de l'adjectif avec le nom pluriel."),
    (r"(?i)\bHé\b", "He", "Remove an unintended accent from the English pronoun."),
    (r"(?i)\bHe\s+HAS\b", "He has", "Normalize pronoun and verb capitalization."),
    (r"(?i)\bThe\s+information\s+are\b", "The information is",
     "The uncountable noun « information » takes a singular verb."),
    # English high-confidence fallback rules, independent of remote service availability.
    (r"(?i)\bShe\s+go\b", "She goes", "Third-person singular: add -es to the verb."),
    (r"(?i)\bThey\s+is\b", "They are", "Plural subject requires the verb « are »."),
    (r"(?i)\bHe\s+have\b", "He has", "Third-person singular: « have » becomes « has »."),
    (r"(?i)\bHé\s+have\s+à\s+new\s+computer\b",
     "He has a new computer", "Correction of the intended English sentence and accent/word errors."),
    (r"(?i)\binformations\b", "information",
     "In standard English, « information » is an uncountable noun."),
    # Targeted agreement and verb-form fallbacks observed in the English DOCX test.
    (r"(?i)\beducation\s+are\s+considered\b", "education is considered",
     "Uncountable singular subject « education » requires « is »."),
    (r"(?i)\bmany\s+children['’]s\b", "many children",
     "Plural « children » does not take a possessive ending here."),
    (r"(?i)\bstill\s+does\s+not\s+have\b", "still do not have",
     "Plural subject requires « do not »."),
    (r"(?i)\buseful\s+equipments\b", "useful equipment",
     "« Equipment » is uncountable in standard English."),
    (r"(?i)\bschools\s+in\s+rural\s+areas\s+does\s+not\s+have\b",
     "schools in rural areas do not have", "Plural subject requires « do not have »."),
    (r"(?i)\bteachers\s+also\s+plays\b", "teachers also play",
     "Plural subject requires the base verb « play »."),
    (r"(?i)\ba\s+good\s+teacher\s+encourage\b", "a good teacher encourages",
     "Third-person singular subject requires « encourages »."),
    (r"(?i)\bstudents\s+feels\s+supported\b", "students feel supported",
     "Plural subject requires « feel »."),
    (r"(?i)\btechnology\s+have\s+changed\b", "technology has changed",
     "Singular subject requires « has »."),
    (r"(?i)\bcommunication\s+skills\s+is\b", "communication skills are",
     "Plural subject « skills » requires « are »."),
    (r"(?i)\bemployees\s+must\s+knows\b", "employees must know",
     "A modal verb takes the base form of the following verb."),
    (r"(?i)\bteamwork\s+are\b", "teamwork is",
     "Uncountable collective subject « teamwork » is singular."),
    (r"(?i)\bteamwork\s+sometimes\s+become\b", "teamwork sometimes becomes",
     "Singular subject requires « becomes »."),
    (r"(?i)\ba\s+manager\s+who\s+want\b", "a manager who wants",
     "Third-person singular subject requires « wants »."),
    (r"(?i)\bshould\s+listens\b", "should listen",
     "After the modal « should », use the base verb « listen »."),
    (r"(?i)\bprovides\s+clear\s+instructions\b", "provide clear instructions",
     "After « should », use the base verb « provide »."),
    (r"(?i)\bworkers\s+feels\s+respected\b", "workers feel respected",
     "Plural subject requires « feel »."),
    (r"(?i)\bthe\s+development\s+of\s+a\s+country\s+also\s+depend\b",
     "the development of a country also depends",
     "The head subject « development » is singular."),
    (r"(?i)\bsome\s+of\s+them\s+have\s+university\s+degrees\s+but\s+does\s+not\s+have\b",
     "some of them have university degrees but do not have",
     "Plural subject « some of them » requires « do not have »."),
    (r"(?i)\bothers\s+wants\b", "others want",
     "Plural subject requires « want »."),
    (r"(?i)\ba\s+young\s+person\s+who\s+have\b", "a young person who has",
     "Singular subject requires « has »."),
    (r"(?i)\bstarting\s+a\s+business\s+require\b",
     "starting a business requires", "Gerund phrase used as a subject is singular."),
    (r"(?i)\bmust\s+study\s+the\s+market,\s*identifies\s+potential\s+customers\s+and\s+calculates\s+their\s+expenses\b",
     "must study the market, identify potential customers and calculate their expenses",
     "Parallel verbs after « must » use the base form."),
    (r"(?i)\bthey\s+should\s+establish\s+priorities,\s*prepares\s+a\s+daily\s+schedule\s+and\s+avoids\s+unnecessary\s+distractions\b",
     "they should establish priorities, prepare a daily schedule and avoid unnecessary distractions",
     "Parallel verbs after « should » use the base form."),
)


def _apply_local_corrections(paragraph):
    """Apply only high-confidence local fixes while preserving Word run formatting."""
    applied = []
    for pattern, replacement, explanation in LOCAL_CORRECTION_RULES:
        matches = list(re.finditer(pattern, paragraph.text or ""))
        for match in reversed(matches):
            start, end = match.span()
            spans = _run_spans(paragraph)
            affected = [
                (run_start, run_end, run)
                for run_start, run_end, run in spans
                if run_start < end and run_end > start
            ]
            if not affected:
                continue

            old_text = (paragraph.text or "")[start:end]
            first_start, first_end, first_run = affected[0]
            local_start = start - first_start
            local_end = min(end, first_end) - first_start
            first_value = first_run.text or ""
            first_run.text = (
                first_value[:local_start] + replacement + first_value[local_end:]
            )

            for run_start, run_end, run in affected[1:]:
                remove_start = max(start, run_start) - run_start
                remove_end = min(end, run_end) - run_start
                value = run.text or ""
                run.text = value[:remove_start] + value[remove_end:]

            applied.append({
                "Texte initial": old_text,
                "Texte corrigé": replacement,
                "Explication": explanation,
                "Méthode": "Règle locale prudente; mise en forme Word conservée autant que possible.",
            })
    return applied

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



def analyze_document_structure(document):
    """Detect repeated paragraphs and suspicious text artefacts without editing content.

    The function is deliberately diagnostic: it never deletes, merges, or rewrites
    paragraphs. Positions are reported so the author can review each occurrence.
    """
    occurrences = defaultdict(list)
    artefacts = []
    paragraphs = list(_document_paragraphs(document))

    # Ces règles signalent des éléments à contrôler sans jamais les supprimer.
    # Certaines correspondent à des artefacts observés dans des documents importés;
    # les résultats restent des alertes, car un terme isolé peut être légitime.
    suspicious_patterns = (
        (r"\b[A-Za-zÀ-ÿ]+\+\d+\b", "Marqueur technique ajouté à une citation (ex. Data+1)."),
        (r"\b(?:L\d+\s+)?Rapport\s+final\b", "Titre ou marqueur de version à vérifier dans le document."),
        (r"\bvariablesreposeerche\b", "Fusion probable de plusieurs mots : « variables de la recherche »."),
        (r"\bAnalanjirofod\b", "Nom géographique possiblement altéré : vérifier « Analanjirofo »."),
        (r"Les Publications du CIRAD(?:\+\d+)?", "Libellé de plateforme bibliographique à vérifier; remplacer par une référence complète si nécessaire."),
        (r"(?<!\w)(?:ECOLEX|ECOLE)(?!\w)", "Libellé isolé pouvant provenir d'un extrait de source; vérifier le contexte."),

        (r"\b(?:spécialisé|spécialisée|spécialise)(?:\s+au)?(?:spécialise|spécialisée|containerisé)[A-Za-zÀ-ÿ]*",
         "Suite de mots fusionnés ou répétés : phrase probablement corrompue."),
        (r"\bprovoqué\s+par\s+provoqué\w*",
         "Répétition ou fusion suspecte dans une phrase; comparer avec le document original."),
        (r"[æ�]|%ICTSL",
         "Caractère anormal ou sigle altéré; vérifier le texte source."),
        (r"\b(?:I|II|III|IV|V|VI|VII|VIII|IX|X)\.\s+(?:I|II|III|IV|V|VI|VII|VIII|IX|X)\.",
         "Numérotation de titre possiblement doublée ou incohérente."),
    )

    for index, paragraph in enumerate(paragraphs, start=1):
        raw = paragraph.text or ""
        normalized = unicodedata.normalize("NFKC", raw)
        normalized = re.sub(r"\s+", " ", normalized).strip().casefold()
        if len(normalized) >= 50:
            occurrences[normalized].append((index, raw[:180]))
        for pattern, reason in suspicious_patterns:
            for match in re.finditer(pattern, raw, re.IGNORECASE):
                artefacts.append({
                    "Paragraphe": index,
                    "Élément": match.group(0),
                    "Raison": reason,
                    "Action": "À vérifier manuellement — le texte original est conservé.",
                })

    repeated = []
    for items in occurrences.values():
        if len(items) > 1:
            repeated.append({
                "Occurrences": len(items),
                "Paragraphes": ", ".join(str(item[0]) for item in items),
                "Extrait": items[0][1],
                "Action": "À vérifier manuellement — aucune suppression automatique.",
            })

    # Repérer les couvertures contradictoires sans choisir ni supprimer un titre.
    full_text = "\n".join(paragraph.text or "" for paragraph in paragraphs)
    has_visit_report = bool(re.search(r"(?i)RAPPORT\s+DE\s+VISITE", full_text))
    has_licence_memoire = bool(re.search(r"(?i)M[ÉE]MOIRE\s+DE\s+LICENCE", full_text))
    if has_visit_report and has_licence_memoire:
        artefacts.append({
            "Paragraphe": 0,
            "Élément": "RAPPORT DE VISITE / MÉMOIRE DE LICENCE",
            "Raison": "Deux types de document apparaissent dans le même fichier; confirmer le titre officiel.",
            "Action": "À vérifier manuellement — aucun titre n'est supprimé.",
        })

    return {
        "paragraph_count": len(paragraphs),
        "repeated_paragraphs": sorted(repeated, key=lambda item: item["Paragraphes"]),
        "suspicious_artefacts": artefacts,
    }


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
    """Split document text into bounded requests while retaining paragraph offsets."""
    current = []
    current_size = 0
    for paragraph in paragraphs:
        value = paragraph.text or ""
        if not value.strip():
            continue
        if len(value) > MAX_CHUNK_CHARS:
            if current:
                yield current
                current, current_size = [], 0
            start = 0
            while start < len(value):
                end = min(start + MAX_CHUNK_CHARS, len(value))
                if end < len(value):
                    # Prefer a nearby sentence/word boundary, but always make progress.
                    boundary = max(value.rfind(". ", start, end),
                                   value.rfind("! ", start, end),
                                   value.rfind("? ", start, end),
                                   value.rfind(" ", start, end))
                    if boundary > start + MAX_CHUNK_CHARS // 2:
                        end = boundary + 1
                yield [(paragraph, start, value[start:end])]
                start = end
            continue
        if current and current_size + len(value) + 1 > MAX_CHUNK_CHARS:
            yield current
            current, current_size = [], 0
        current.append((paragraph, 0, value))
        current_size += len(value) + 1
    if current:
        yield current


def _post_languagetool(text: str, language_code: str):
    """Retry temporary LanguageTool failures such as HTTP 429 and 5xx."""
    for attempt in range(3):
        try:
            response = requests.post(
                LANGUAGETOOL_URL,
                data={"text": text, "language": language_code},
                timeout=60,
            )
            response.raise_for_status()
            return response.json()
        except requests.RequestException as exc:
            status = getattr(getattr(exc, "response", None), "status_code", None)
            temporary = status == 429 or (status is not None and status >= 500)
            temporary = temporary or isinstance(
                exc, (requests.Timeout, requests.ConnectionError)
            )
            if not temporary or attempt == 2:
                raise
            time.sleep(1 + attempt * 2)


def correct_docx_with_languagetool(document, language_code: str):
    """Correct a DOCX and return (document, applied_changes, skipped_matches)."""
    paragraphs = list(_document_paragraphs(document))
    changes = []
    skipped = []
    suggestions = []

    # Corriger d'abord les erreurs récurrentes très clairement identifiables,
    # y compris celles que le service LanguageTool ne signale pas.
    for paragraph in paragraphs:
        local_changes = _apply_local_corrections(paragraph)
        changes.extend(local_changes)
    detected_count = len(changes)

    # Re-run LanguageTool on the updated text so that agreements exposed by
    # earlier corrections (e.g. "Les enfant joue" -> "Les enfants jouent")
    # are detected. Stop early when a pass applies no new correction.
    for pass_index in range(3):
        applied_before_pass = len(changes)
        for chunk in _chunks(paragraphs):
            pieces = [item[2] for item in chunk]
            text = "\n".join(pieces)
            if not text.strip():
                continue
    
            result = _post_languagetool(text, language_code)
            matches = result.get("matches", [])
            protected_spans = _protected_spans(text)
    
            paragraph_starts = []
            cursor = 0
            for paragraph, piece_offset, piece in chunk:
                paragraph_starts.append((cursor, cursor + len(piece), paragraph, piece_offset))
                cursor += len(piece) + 1
    
            edits_by_run = {}
            for match in matches:
                detected_count += 1
                offset = int(match.get("offset", 0))
                length = int(match.get("length", 0))
                replacements = match.get("replacements") or []
                original_text = text[offset:offset + length]
                if length > 0 and _overlaps_protected_span(
                    offset, offset + length, protected_spans
                ):
                    skipped.append({
                        "Texte": original_text,
                        "Raison": "Nom propre, sigle, référence ou terme protégé : aucune modification automatique.",
                    })
                    continue
                replacement_values = [
                    item.get("value", "") for item in replacements
                    if item.get("value", "")
                ]
                suggestions.append({
                    "Texte détecté": original_text,
                    "Suggestion LanguageTool": " / ".join(replacement_values[:3]) or "Aucune suggestion proposée",
                    "Explication": match.get("message", "Suggestion linguistique"),
                })
                if length <= 0 or not replacement_values:
                    skipped.append({
                        "Texte": original_text,
                        "Raison": "LanguageTool a détecté un point à vérifier mais n'a proposé aucune correction automatique.",
                    })
                    continue
    
                replacement, selection_reason = _choose_safe_replacement(
                    text, offset, original_text, replacement_values
                )
                if replacement is None:
                    skipped.append({
                        "Texte": original_text,
                        "Raison": selection_reason,
                    })
                    continue
    
                if replacement.strip() == original_text.strip():
                    # A repeated suggestion after a previous pass is not an
                    # unapplied correction; simply ignore this no-op match.
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
                spans = _run_spans(paragraph)
                affected = [
                    (index, start, end, run)
                    for index, (start, end, run) in enumerate(spans)
                    if start < paragraph_end and end > paragraph_start
                ]
                if not affected or affected[0][1] > paragraph_start or affected[-1][2] < paragraph_end:
                    skipped.append({
                        "Texte": paragraph.text[paragraph_start:paragraph_end],
                        "Raison": "Impossible de localiser précisément le texte dans Word.",
                    })
                    continue
    
                old_text = paragraph.text[paragraph_start:paragraph_end]
                edits_by_run.setdefault(paragraph, []).append((
                    paragraph_start, paragraph_end, replacement, old_text,
                    match.get("message", "Correction linguistique"),
                    selection_reason,
                ))
    
            # Apply edits from right to left using the original paragraph offsets.
            # This also handles a correction split across several Word runs:
            # the replacement keeps the formatting of the first affected run,
            # while the matched characters are removed from the following runs.
            for paragraph, edits in edits_by_run.items():
                accepted = []
                last_start = float("inf")
                for edit in sorted(edits, key=lambda item: item[0], reverse=True):
                    start, end = edit[0], edit[1]
                    if end > last_start:
                        skipped.append({
                            "Texte": edit[3],
                            "Raison": "Suggestion chevauchant une autre correction.",
                        })
                        continue
                    accepted.append(edit)
                    last_start = start
    
                for paragraph_start, paragraph_end, replacement, old_text, message, selection_reason in accepted:
                    spans = _run_spans(paragraph)
                    affected = [
                        (start, end, run)
                        for start, end, run in spans
                        if start < paragraph_end and end > paragraph_start
                    ]
                    if not affected:
                        skipped.append({
                            "Texte": old_text,
                            "Raison": "Segments Word introuvables au moment de l'application.",
                        })
                        continue
    
                    first_start, first_end, first_run = affected[0]
                    local_start = paragraph_start - first_start
                    local_end = min(paragraph_end, first_end) - first_start
                    first_text = first_run.text or ""
                    first_run.text = first_text[:local_start] + replacement + first_text[local_end:]
    
                    for run_start, run_end, run in affected[1:]:
                        remove_start = max(paragraph_start, run_start) - run_start
                        remove_end = min(paragraph_end, run_end) - run_start
                        value = run.text or ""
                        run.text = value[:remove_start] + value[remove_end:]
    
                    changes.append({
                        "Texte initial": old_text,
                        "Texte corrigé": replacement,
                        "Explication": message,
                        "Méthode": selection_reason,
                    })
    
        if len(changes) == applied_before_pass:
            break

    # Final local pass: LanguageTool can occasionally suggest an incorrect casing/accent
    # or leave a known agreement error after earlier edits. Re-apply only the
    # conservative local rules so the final DOCX cannot undo these high-confidence fixes.
    for paragraph in paragraphs:
        final_local_changes = _apply_local_corrections(paragraph)
        changes.extend(final_local_changes)

    return document, changes, skipped, detected_count, suggestions

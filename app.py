import io, os, re, tempfile, subprocess, shutil
from pathlib import Path
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import requests
from correction_engine import correct_docx_with_languagetool, analyze_document_structure
from export_engine import add_page_numbers, apply_academic_heading_styles, build_pdf_from_docx

APP_NAME = "Mémoire Correcteur Pro"
st.set_page_config(page_title=APP_NAME, page_icon="📘", layout="wide")
st.title("📘 CORRECTGLOSS")
st.caption("Correction • reformulation académique • mise en page • contrôle des citations")

def log_app_event(event_type):
    url = os.getenv("SUPABASE_URL", "").rstrip("/")
    anon = os.getenv("SUPABASE_ANON_KEY", "")
    if not url or not anon:
        return
    try:
        requests.post(f"{url}/rest/v1/app_events",
            headers={"apikey": anon, "Authorization": f"Bearer {anon}",
                     "Content-Type": "application/json", "Prefer": "return=minimal"},
            json={"event_name": event_type, "app_version": "0.4"}, timeout=5).raise_for_status()
    except Exception:
        pass

if "cg_session_logged" not in st.session_state:
    log_app_event("app_open")
    st.session_state["cg_session_logged"] = True

def add_word_toc(document):
    """Replace a clearly labelled initial manual TOC with a real Word TOC field."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    paragraphs = list(document.paragraphs)
    if any("TOC" in paragraph._p.xml and "instrText" in paragraph._p.xml
           for paragraph in paragraphs):
        return document

    # Remove only an initial section explicitly titled "Table des matières",
    # and only when a clear project-body heading marks where the contents end.
    toc_index = next(
        (i for i, p in enumerate(paragraphs)
         if p.text.strip().casefold() in ("table des matières", "table des matieres", "sommaire")),
        None,
    )
    if toc_index is not None:
        body_index = next(
            (i for i in range(toc_index + 1, min(len(paragraphs), toc_index + 250))
             if paragraphs[i].text.strip().upper().startswith("PROJET DE RECHERCHE")),
            None,
        )
        if body_index is not None:
            for paragraph in paragraphs[toc_index:body_index]:
                parent = paragraph._p.getparent()
                if parent is not None:
                    parent.remove(paragraph._p)

    title = document.add_paragraph("TABLE DES MATIÈRES")
    # Some uploaded DOCX templates omit the built-in "Title" paragraph style.
    # Use the guaranteed Normal style and apply title formatting directly.
    try:
        title.style = document.styles["Title"]
    except KeyError:
        title.style = document.styles["Normal"]
    title.alignment = 1  # centered
    for title_run in title.runs:
        title_run.bold = True
        title_run.font.size = Pt(14)
    field_paragraph = document.add_paragraph()
    run = field_paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = ' TOC \\o "1-3" \\h \\z \\u '
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for element in (begin, instruction, separate, end):
        run._r.append(element)

    body = document._element.body
    body.remove(title._p)
    body.remove(field_paragraph._p)
    body.insert(0, title._p)
    body.insert(1, field_paragraph._p)

    settings = document.settings._element
    update_fields = settings.find(qn("w:updateFields"))
    if update_fields is None:
        update_fields = OxmlElement("w:updateFields")
        settings.append(update_fields)
    update_fields.set(qn("w:val"), "true")
    return document


with st.sidebar:
    st.header("Paramètres")
    mode = st.radio("Mode de correction", ["Hors ligne", "En ligne (LanguageTool)"])
    language = st.selectbox("Langue", ["Français", "Anglais", "Malagasy"])
    font = st.selectbox("Police", ["Times New Roman", "Arial", "Calibri"])
    font_size = st.number_input("Taille (pt)", 10, 14, 12)
    line_spacing = st.selectbox("Interligne", ["1.0", "1.15", "1.5", "2.0"], index=2)
    margin_cm = st.number_input("Marges (cm)", 1.5, 4.0, 2.5, 0.25)
    justify = st.checkbox("Justifier le texte", True)


st.divider()
st.header("📝 Préparer une rédaction")
st.caption("Préparez un plan ou un brouillon structuré. Cette fonction fonctionne sans API payante.")
with st.form("cg_prepare_writing_form"):
    writing_mode = st.radio(
        "Que souhaitez-vous préparer ?",
        ["Rédaction complète (brouillon à compléter)", "Plan détaillé et accompagnement"],
        help="La rédaction complète crée une trame structurée avec des zones à compléter. Le plan détaillé vous guide étape par étape."
    )
    writing_subject = st.text_input("Sujet / thème", placeholder="Ex. : La compétitivité du girofle malgache à l'exportation")
    writing_type = st.selectbox(
        "Type de rédaction",
        ["Dissertation", "Exposé", "Rapport", "Mémoire / Projet de recherche", "Lettre", "Article"],
    )
    writing_level = st.selectbox(
        "Niveau", ["Collège", "Lycée", "Licence", "Master", "Professionnel"]
    )
    writing_language = st.selectbox("Langue de rédaction", ["Français", "Anglais"])
    writing_instructions = st.text_area(
        "Consignes particulières (facultatif)",
        placeholder="Ex. : parler du contexte malgache, citer des exemples, préciser la période étudiée..."
    )
    writing_submit = st.form_submit_button("Préparer ma rédaction", type="primary")

def cg_build_writing_outline(subject, kind, level, lang, instructions, full_draft):
    is_fr = lang == "Français"
    labels = {
        "fr": {
            "intro": "INTRODUCTION",
            "context": "Contexte et justification du sujet",
            "problem": "Problématique",
            "questions": "Questions de recherche / objectifs",
            "plan": "Annonce du plan",
            "development": "DÉVELOPPEMENT",
            "part1": "I. Première partie : notions, définitions et cadre général",
            "part2": "II. Deuxième partie : analyse, arguments et exemples",
            "part3": "III. Troisième partie : résultats, discussion ou propositions",
            "conclusion": "CONCLUSION",
            "summary": "Synthèse des idées principales",
            "answer": "Réponse à la problématique",
            "limits": "Limites et perspectives",
            "sources": "SOURCES ET RÉFÉRENCES",
            "source_note": "Insérez uniquement les ouvrages, articles, rapports et sites réellement consultés.",
            "fill": "[À compléter avec vos informations, vos données et vos sources vérifiées.]",
            "guide": "Questions pour vous guider",
            "g1": "Quel est le contexte du sujet et pourquoi est-il important ?",
            "g2": "Quel problème précis voulez-vous étudier ou expliquer ?",
            "g3": "Quels faits, données ou exemples vérifiables soutiennent votre analyse ?",
            "g4": "Que montrent vos résultats et quelles conclusions pouvez-vous en tirer ?",
            "instructions": "Consignes fournies",
        },
        "en": {
            "intro": "INTRODUCTION",
            "context": "Background and rationale",
            "problem": "Problem statement",
            "questions": "Research questions / objectives",
            "plan": "Outline of the paper",
            "development": "MAIN BODY",
            "part1": "I. Part One: concepts, definitions and general framework",
            "part2": "II. Part Two: analysis, arguments and examples",
            "part3": "III. Part Three: findings, discussion or recommendations",
            "conclusion": "CONCLUSION",
            "summary": "Summary of the main ideas",
            "answer": "Answer to the problem statement",
            "limits": "Limitations and future perspectives",
            "sources": "SOURCES AND REFERENCES",
            "source_note": "Include only books, articles, reports and websites you actually consulted.",
            "fill": "[Complete with your own information, data and verified sources.]",
            "guide": "Guiding questions",
            "g1": "What is the context of the topic, and why does it matter?",
            "g2": "What specific issue do you want to study or explain?",
            "g3": "Which verifiable facts, data or examples support your analysis?",
            "g4": "What do your findings show, and what conclusions can you draw?",
            "instructions": "Your instructions",
        },
    }
    t = labels["fr" if is_fr else "en"]
    if kind == "Lettre":
        sections = [
            ("[Lieu et date]", t["fill"]),
            ("[Destinataire]", t["fill"]),
            ("Objet : " + subject, t["fill"]),
            (t["intro"], t["fill"]),
            ("Développement", t["fill"]),
            ("Formule de politesse et signature", t["fill"]),
        ]
    elif kind == "Article":
        sections = [
            (t["intro"], t["context"] + "\n" + t["fill"]),
            (t["development"], t["part1"] + "\n" + t["fill"] + "\n\n" + t["part2"] + "\n" + t["fill"]),
            (t["conclusion"], t["summary"] + "\n" + t["fill"]),
            (t["sources"], t["source_note"]),
        ]
    else:
        sections = [
            (t["intro"], t["context"] + "\n" + t["fill"] + "\n\n" + t["problem"] + "\n" + t["fill"] + "\n\n" + t["questions"] + "\n" + t["fill"] + "\n\n" + t["plan"] + "\n" + t["fill"]),
            (t["development"], t["part1"] + "\n" + t["fill"] + "\n\n" + t["part2"] + "\n" + t["fill"] + "\n\n" + t["part3"] + "\n" + t["fill"]),
            (t["conclusion"], t["summary"] + "\n" + t["fill"] + "\n\n" + t["answer"] + "\n" + t["fill"] + "\n\n" + t["limits"] + "\n" + t["fill"]),
            (t["sources"], t["source_note"]),
        ]
    lines = [
        ("TITLE", subject),
        ("Type", kind),
        ("Level", level),
        ("Language", lang),
    ]
    if instructions.strip():
        lines.append((t["instructions"], instructions.strip()))
    if full_draft:
        lines.append((t["intro"], t["fill"]))
    for heading, body in sections:
        lines.append((heading, body))
    if not full_draft:
        lines.extend([
            (t["guide"], "\n".join([t["g1"], t["g2"], t["g3"], t["g4"]])),
        ])
    return lines

if writing_submit:
    if not writing_subject.strip():
        st.error("Veuillez saisir le sujet de votre rédaction.")
    else:
        prepared_lines = cg_build_writing_outline(
            writing_subject.strip(), writing_type, writing_level, writing_language,
            writing_instructions, writing_mode.startswith("Rédaction complète")
        )
        prepared_doc = Document()
        prepared_doc.add_heading(writing_subject.strip(), 0)
        prepared_doc.add_paragraph(
            f"Type : {writing_type} | Niveau : {writing_level} | Langue : {writing_language}"
        )
        prepared_text_parts = [writing_subject.strip()]
        for heading, body in prepared_lines[4:]:
            prepared_doc.add_heading(heading, level=1)
            for paragraph_text in body.split("\n"):
                prepared_doc.add_paragraph(paragraph_text)
            prepared_text_parts.extend([heading, body])
        prepared_buffer = io.BytesIO()
        prepared_doc.save(prepared_buffer)
        st.session_state["cg_prepared_writing_docx"] = prepared_buffer.getvalue()
        st.session_state["cg_prepared_writing_text"] = "\n\n".join(prepared_text_parts)
        st.session_state["cg_prepared_writing_name"] = re.sub(
            r"[^A-Za-z0-9_-]+", "_", writing_subject.strip()
        )[:50] or "redaction"
        st.success(
            "Votre plan structuré est prêt." if writing_mode.startswith("Plan détaillé")
            else "Votre brouillon structuré est prêt. Complétez les passages entre crochets avec vos propres informations et sources."
        )
        st.info("Version gratuite : CORRECTGLOSS prépare une structure et des indications. Il n'invente pas de données, de citations ou de références bibliographiques.")

if "cg_prepared_writing_docx" in st.session_state:
    st.download_button(
        "Télécharger la rédaction préparée (Word)",
        st.session_state["cg_prepared_writing_docx"],
        file_name=f"correctgloss_{st.session_state.get('cg_prepared_writing_name', 'redaction')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    with st.expander("Afficher le contenu préparé"):
        st.text_area("Contenu", st.session_state["cg_prepared_writing_text"], height=300)

uploaded = st.file_uploader("Importer un mémoire Word (.docx)", type=["docx"])
if not uploaded:
    st.info("Importez un document Word pour lancer les contrôles.")
    st.stop()

try:
    original_bytes = uploaded.getvalue()
    doc = Document(io.BytesIO(original_bytes))
except Exception as e:
    st.error(f"Impossible d'ouvrir le document : {e}")
    st.stop()

paragraphs = [p.text for p in doc.paragraphs]
text = "\n\n".join(p for p in paragraphs if p.strip())
st.success(f"Document chargé : {len(paragraphs)} paragraphes • {len(doc.tables)} tableaux")

structure_report = analyze_document_structure(doc)
with st.expander("Contrôle de la structure du document", expanded=True):
    st.caption("Ce contrôle signale les répétitions et artefacts possibles. Il ne supprime et ne modifie aucun contenu.")
    st.metric("Paragraphes examinés", structure_report["paragraph_count"])
    repeated = structure_report["repeated_paragraphs"]
    artefacts = structure_report["suspicious_artefacts"]
    if repeated:
        st.warning(f"{len(repeated)} groupe(s) de paragraphes répétés à vérifier.")
        st.dataframe(repeated, use_container_width=True, hide_index=True)
    else:
        st.success("Aucun paragraphe long identique détecté.")
    if artefacts:
        st.warning(f"{len(artefacts)} élément(s) pouvant être des artefacts de copie ou de génération.")
        st.dataframe(artefacts, use_container_width=True, hide_index=True)
    else:
        st.info("Aucun artefact correspondant aux règles de détection n'a été repéré.")

st.text_area("Aperçu", "\n\n".join([p for p in paragraphs if p.strip()][:8]), height=220)

with st.expander("Correction manuelle du texte", expanded=False):
    manual_source_bytes = st.session_state.get("cg_corrected_docx", original_bytes)
    manual_document = Document(io.BytesIO(manual_source_bytes))
    manual_paragraphs = [p for p in manual_document.paragraphs if p.text.strip()]
    manual_initial_text = "\n\n".join(p.text for p in manual_paragraphs)
    manual_version = "corrige" if "cg_corrected_docx" in st.session_state else "original"
    manual_text = st.text_area(
        "Modifiez directement le texte (gardez le même nombre de paragraphes séparés par une ligne vide).",
        value=manual_initial_text,
        height=300,
        key=f"cg_manual_text_{uploaded.name}_{len(original_bytes)}_{manual_version}",
    )
    st.caption("La correction manuelle concerne les paragraphes du corps du document. Relisez la mise en forme après application.")
    if st.button("Appliquer les corrections manuelles"):
        edited_paragraphs = manual_text.split("\n\n")
        if len(edited_paragraphs) != len(manual_paragraphs):
            st.error(
                f"Le nombre de paragraphes doit rester identique : "
                f"{len(manual_paragraphs)} attendu(s), {len(edited_paragraphs)} reçu(s)."
            )
        else:
            for paragraph, edited_value in zip(manual_paragraphs, edited_paragraphs):
                paragraph.text = edited_value
            manual_output = io.BytesIO()
            manual_document.save(manual_output)
            st.session_state["cg_corrected_docx"] = manual_output.getvalue()
            st.session_state.pop("cg_docx", None)
            st.session_state.pop("cg_pdf", None)
            log_app_event("manual_correction_applied")
            st.success("Corrections manuelles appliquées. Préparez ensuite les fichiers Word et PDF.")

st.subheader("Correction automatique du document Word")
st.caption("La correction utilise LanguageTool en ligne et peut corriger un mot même s’il est réparti sur plusieurs segments de mise en forme Word.")
if st.button("Corriger automatiquement le document", type="primary"):
    if language == "Malagasy":
        st.warning("La correction automatique LanguageTool n'est pas disponible pour le malagasy dans cette version.")
    else:
        try:
            correction_doc = Document(io.BytesIO(original_bytes))
            language_code = "fr" if language == "Français" else "en-US"
            correction_doc, applied_changes, skipped_matches, detected_count, detected_suggestions = correct_docx_with_languagetool(
                correction_doc, language_code
            )
            # Refaire le diagnostic sur le document réellement corrigé,
            # afin de signaler les artefacts éventuellement laissés après correction.
            corrected_structure_report = analyze_document_structure(correction_doc)
            corrected_output = io.BytesIO()
            correction_doc.save(corrected_output)
            st.session_state["cg_corrected_docx"] = corrected_output.getvalue()
            st.session_state["cg_correction_changes"] = applied_changes
            st.session_state["cg_correction_skipped"] = skipped_matches
            st.session_state["cg_correction_detected"] = detected_count
            st.session_state["cg_detected_suggestions"] = detected_suggestions
            st.session_state["cg_corrected_structure_report"] = corrected_structure_report
            st.success("Correction automatique terminée.")
        except requests.RequestException as e:
            st.error(f"LanguageTool est momentanément indisponible : {e}")
        except Exception as e:
            st.error(f"Erreur pendant la correction : {e}")

if "cg_corrected_docx" in st.session_state:
    st.download_button(
        "Télécharger le document Word corrigé",
        st.session_state["cg_corrected_docx"],
        file_name="correctgloss_corrige.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    detected_count = st.session_state.get("cg_correction_detected", 0)
    applied_count = len(st.session_state.get("cg_correction_changes", []))
    not_applied_count = max(0, detected_count - applied_count)
    metric_cols = st.columns(3)
    metric_cols[0].metric("Erreurs détectées", detected_count)
    metric_cols[1].metric("Corrections appliquées", applied_count)
    metric_cols[2].metric("Non appliquées", not_applied_count)

    corrected_preview_doc = Document(io.BytesIO(st.session_state["cg_corrected_docx"]))
    corrected_preview_text = "\n\n".join(
        p.text for p in corrected_preview_doc.paragraphs if p.text.strip()
    )
    st.text_area("Aperçu du document corrigé", corrected_preview_text, height=220)
    corrected_report = st.session_state.get("cg_corrected_structure_report")
    if corrected_report:
        with st.expander("Anomalies restantes après correction automatique", expanded=True):
            st.caption("Ces éléments sont signalés pour vérification. CORRECTGLOSS ne supprime pas automatiquement les paragraphes, titres ou références.")
            if corrected_report["repeated_paragraphs"]:
                st.warning("Des paragraphes identiques restent présents.")
                st.dataframe(corrected_report["repeated_paragraphs"], use_container_width=True, hide_index=True)
            if corrected_report["suspicious_artefacts"]:
                st.warning("Des fragments de texte ou incohérences de structure restent à vérifier.")
                st.dataframe(corrected_report["suspicious_artefacts"], use_container_width=True, hide_index=True)
            if not corrected_report["repeated_paragraphs"] and not corrected_report["suspicious_artefacts"]:
                st.success("Aucune anomalie de structure correspondant aux règles de contrôle n'a été détectée.")
    if st.session_state.get("cg_detected_suggestions"):
        st.write("Suggestions détectées par LanguageTool")
        st.dataframe(st.session_state["cg_detected_suggestions"], use_container_width=True, hide_index=True)
    if st.session_state.get("cg_correction_changes"):
        st.write("Modifications appliquées")
        st.dataframe(st.session_state["cg_correction_changes"], use_container_width=True, hide_index=True)
    else:
        st.info("Aucune correction automatique proposée. Cela ne garantit pas l'absence d'erreurs.")
    if st.session_state.get("cg_correction_skipped"):
        st.warning(f"{len(st.session_state['cg_correction_skipped'])} suggestion(s) non appliquée(s) pour préserver le document.")
        st.dataframe(st.session_state["cg_correction_skipped"], use_container_width=True, hide_index=True)

def offline_suggestions(content):
    rules = [
        (r"\bmalgré que\b", "bien que / malgré le fait que", "Formulation à vérifier."),
        (r"\bau jour d'aujourd'hui\b", "aujourd'hui / à ce jour", "Redondance."),
        (r"\btrès important\b", "essentiel / majeur / déterminant", "Préciser le sens."),
        (r"\bon peut voir que\b", "les résultats montrent que", "Style académique."),
        (r"\bje pense que\b", "les données suggèrent que", "Éviter une opinion non justifiée."),
    ]
    found = []
    for pattern, suggestion, note in rules:
        for m in re.finditer(pattern, content, re.I):
            found.append({"motif": m.group(), "suggestion": suggestion, "explication": note})
    return found

if st.button("Analyser le mémoire", type="primary"):
    results = offline_suggestions(text)
    if mode == "En ligne (LanguageTool)" and language != "Malagasy":
        code = "fr" if language == "Français" else "en-US"
        try:
            resp = requests.post("https://api.languagetool.org/v2/check",
                                 data={"text": text[:20000], "language": code}, timeout=25)
            resp.raise_for_status()
            for item in resp.json().get("matches", []):
                start, length = item.get("offset", 0), item.get("length", 0)
                replacements = [x.get("value", "") for x in item.get("replacements", [])[:3]]
                results.append({"motif": text[start:start+length],
                    "suggestion": " / ".join(replacements) or "À vérifier",
                    "explication": item.get("message", "LanguageTool")})
        except Exception as e:
            st.warning(f"LanguageTool indisponible : {e}")
    elif language == "Malagasy" and mode != "Hors ligne":
        st.warning("LanguageTool n'est pas activé pour le malagasy dans cette version.")
    st.session_state["cg_results"] = results
    log_app_event("document_analyzed")

if "cg_results" in st.session_state:
    st.subheader("Suggestions")
    st.write(f"{len(st.session_state['cg_results'])} suggestion(s)")
    if st.session_state["cg_results"]:
        st.dataframe(st.session_state["cg_results"], use_container_width=True, hide_index=True)
    else:
        st.info("Aucune suggestion automatique trouvée. Cela ne garantit pas l'absence d'erreurs.")

st.subheader("Reformulation académique (facultative)")
rewrite_text = st.text_area("Texte à reformuler")
if st.button("Reformuler avec l'IA") and rewrite_text.strip():
    key = os.getenv("OPENAI_API_KEY", "")
    if not key:
        st.error("OPENAI_API_KEY n'est pas configurée dans les secrets.")
    else:
        try:
            from openai import OpenAI
            client = OpenAI(api_key=key)
            prompt = "Reformule en français académique sans inventer de faits ou de sources. Conserve le sens:\n\n" + rewrite_text
            result = client.responses.create(model=os.getenv("OPENAI_MODEL", "gpt-4.1-mini"), input=prompt)
            st.text_area("Proposition à vérifier", result.output_text, height=220)
            log_app_event("ai_rewrite")
        except Exception as e:
            st.error(f"Erreur IA : {e}")

st.subheader("Export Word et PDF")
st.caption("Les exports utilisent le document corrigé automatiquement lorsqu'il est disponible, sinon le document original.")
if st.button("Préparer les fichiers Word et PDF", type="primary"):
    source_bytes = st.session_state.get("cg_corrected_docx", original_bytes)
    formatted = Document(io.BytesIO(source_bytes))
    for section in formatted.sections:
        section.top_margin = section.bottom_margin = Inches(margin_cm / 2.54)
        section.left_margin = section.right_margin = Inches(margin_cm / 2.54)
    for para in formatted.paragraphs:
        if justify and not (para.style and para.style.name.lower().startswith("heading")):
            para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        para.paragraph_format.line_spacing = float(line_spacing)
        para.paragraph_format.space_after = Pt(6)
        for run in para.runs:
            run.font.name = font
            run.font.size = Pt(int(font_size))
    apply_academic_heading_styles(formatted)
    add_page_numbers(formatted)
    formatted = add_word_toc(formatted)
    word_output = io.BytesIO()
    formatted.save(word_output)
    docx_bytes = word_output.getvalue()
    st.session_state["cg_docx"] = docx_bytes
    try:
        st.session_state["cg_pdf"] = build_pdf_from_docx(
            docx_bytes, font=font, font_size=font_size,
            line_spacing=line_spacing, margin_cm=margin_cm,
        )
        log_app_event("docx_pdf_export_prepared")
        st.success("Les fichiers Word et PDF sont prêts.")
    except Exception as e:
        st.session_state.pop("cg_pdf", None)
        st.error(f"Le fichier Word est prêt, mais la création du PDF a échoué : {e}")

if "cg_docx" in st.session_state:
    st.download_button("Télécharger le document Word (.docx)", st.session_state["cg_docx"],
        file_name="correctgloss_memoire_corrige.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
if "cg_pdf" in st.session_state:
    st.download_button("Télécharger le document PDF", st.session_state["cg_pdf"],
        file_name="correctgloss_memoire_corrige.pdf",
        mime="application/pdf")

st.caption("Les suggestions sont indicatives. Vérifiez toujours le document final et conservez une copie originale.")

import io, os, re, tempfile, subprocess, shutil
from pathlib import Path
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import requests
from correction_engine import correct_docx_with_languagetool, analyze_document_structure
from export_engine import add_page_numbers, build_pdf_from_docx

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
    """Insert a real Word TOC field without a visible placeholder."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    # Avoid adding a second dynamic table when the source already has one.
    if any("TOC" in paragraph._p.xml and "instrText" in paragraph._p.xml
           for paragraph in document.paragraphs):
        return document

    paragraph = document.add_paragraph()
    run = paragraph.add_run()
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
    body.remove(paragraph._p)
    body.insert(0, paragraph._p)

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
            corrected_output = io.BytesIO()
            correction_doc.save(corrected_output)
            st.session_state["cg_corrected_docx"] = corrected_output.getvalue()
            st.session_state["cg_correction_changes"] = applied_changes
            st.session_state["cg_correction_skipped"] = skipped_matches
            st.session_state["cg_correction_detected"] = detected_count
            st.session_state["cg_detected_suggestions"] = detected_suggestions
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

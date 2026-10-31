import io, os, re, tempfile, subprocess, shutil
from pathlib import Path
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import requests

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


def secret_value(name, default=""):
    try:
        return st.secrets.get(name, os.getenv(name, default))
    except Exception:
        return os.getenv(name, default)

def show_admin_dashboard():
    st.header("📊 Dashboard administrateur")
    password = secret_value("DASHBOARD_PASSWORD")
    if not password:
        st.error("Configurez DASHBOARD_PASSWORD dans les secrets Streamlit.")
        return
    entered = st.text_input("Mot de passe administrateur", type="password")
    if not entered:
        st.info("Entrez le mot de passe pour afficher les statistiques.")
        return
    if entered != password:
        st.error("Mot de passe incorrect.")
        return
    url = secret_value("SUPABASE_URL").rstrip("/")
    service_key = secret_value("SUPABASE_SERVICE_ROLE_KEY")
    if not url or not service_key:
        st.warning("Ajoutez SUPABASE_URL et SUPABASE_SERVICE_ROLE_KEY dans les secrets Streamlit.")
        return
    try:
        response = requests.get(
            f"{url}/rest/v1/app_events",
            headers={"apikey": service_key, "Authorization": f"Bearer {service_key}"},
            params={"select": "event_name,app_version,created_at", "order": "created_at.desc", "limit": "5000"},
            timeout=20)
        response.raise_for_status()
        events = response.json()
        counts = {}
        for event in events:
            name = event.get("event_name", "")
            counts[name] = counts.get(name, 0) + 1
        cols = st.columns(4)
        for col, label, key in zip(cols,
            ["Ouvertures", "Analyses", "Reformulations IA", "Exports DOCX"],
            ["app_open", "document_analyzed", "ai_rewrite", "docx_export_prepared"]):
            col.metric(label, counts.get(key, 0))
        if events:
            import pandas as pd
            df = pd.DataFrame(events)
            df["created_at"] = pd.to_datetime(df["created_at"], errors="coerce", utc=True)
            daily = df.dropna(subset=["created_at"]).assign(jour=lambda x: x["created_at"].dt.strftime("%Y-%m-%d"))
            daily = daily.groupby(["jour", "event_name"]).size().unstack(fill_value=0)
            st.subheader("Activité quotidienne")
            st.line_chart(daily)
            st.subheader("Événements récents")
            st.dataframe(df, use_container_width=True, hide_index=True)
        else:
            st.info("Aucun événement enregistré pour le moment.")
    except Exception as e:
        st.error(f"Impossible de charger les statistiques : {e}")

def detect_and_style_headings(document):
    for para in document.paragraphs:
        value = para.text.strip()
        if not value:
            continue
        upper = value.upper()
        if re.match(r"^(CHAPITRE|PARTIE)\\s+([IVXLCDM]+|[0-9]+)", upper) or upper in {
            "INTRODUCTION", "CONCLUSION", "BIBLIOGRAPHIE", "SITOGRAPHIE", "ANNEXES"}:
            para.style = document.styles["Heading 1"]
        elif re.match(r"^\\d+\\.\\d+\\.\\d+\\s+", value):
            para.style = document.styles["Heading 3"]
        elif re.match(r"^\\d+\\.\\d+\\s+", value):
            para.style = document.styles["Heading 2"]
        elif re.match(r"^\\d+\\.\\s+", value):
            para.style = document.styles["Heading 1"]
    return document

def add_word_toc(document):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    p = document.add_paragraph()
    run = p.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = ' TOC \\\\o "1-3" \\\\h \\\\z \\\\u '
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text_run = OxmlElement("w:r")
    text_node = OxmlElement("w:t")
    text_node.text = "Ouvrez dans Word puis mettez à jour la table."
    text_run.append(text_node)
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.append(begin); run._r.append(instr); run._r.append(separate)
    p._p.append(text_run); run._r.append(end)
    body = document._element.body
    body.remove(p._p)
    body.insert(0, p._p)
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

if page == "Dashboard administrateur":
    show_admin_dashboard()
    st.stop()

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
st.text_area("Aperçu", "\n\n".join([p for p in paragraphs if p.strip()][:8]), height=220)

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

st.subheader("Export et mise en page")
if st.button("Préparer le DOCX avec table des matières", type="primary"):
    formatted = Document(io.BytesIO(original_bytes))
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
    formatted = detect_and_style_headings(formatted)
    formatted = add_word_toc(formatted)
    output = io.BytesIO()
    formatted.save(output)
    st.session_state["cg_docx"] = output.getvalue()
    log_app_event("docx_export_prepared")
    st.success("DOCX préparé. Dans Word, clic droit sur la table des matières > Mettre à jour les champs > Toute la table.")

if "cg_docx" in st.session_state:
    st.download_button("Télécharger le DOCX", st.session_state["cg_docx"],
        file_name="correctgloss_memoire.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

st.caption("Les suggestions sont indicatives. Conservez toujours une copie originale du mémoire.")

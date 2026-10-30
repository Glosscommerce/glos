import requests
import io, os, re, json, tempfile, subprocess, shutil, hmac
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from pathlib import Path
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
import requests

def log_event(event_name):("app_opened")("docx_export_prepared")
APP_NAME = "CORRECTGLOSS"

def log_event(event_name):
    """Enregistre uniquement le type d'événement, jamais le texte du document."""
    try:
        supabase_url = str(st.secrets.get("SUPABASE_URL", "")).strip().rstrip("/")
        supabase_key = str(st.secrets.get("SUPABASE_ANON_KEY", "")).strip()
        if not supabase_url or not supabase_key:
            return
        response = requests.post(
            f"{supabase_url}/rest/v1/app_events",
            headers={
                "apikey": supabase_key,
                "Authorization": f"Bearer {supabase_key}",
                "Content-Type": "application/json",
                "Prefer": "return=minimal",
            },
            json={"event_name": event_name, "app_version": "0.4"},
            timeout=8
        )
        response.raise_for_status()
    except Exception:
        pass
        passAPP_NAME = "CORRECTGLOSS""CORRECTGLOSS"
APP_DIR = Path(__file__).resolve().parent
LOGO_PATH = APP_DIR / "correctgloss_icon.png"
BRAND_PATH = APP_DIR / "correctgloss_brand.png"
st.set_page_config(
    page_title="CORRECTGLOSS | Correcteur de mémoire",
    page_icon=str(LOGO_PATH) if LOGO_PATH.exists() else "✅",
    layout="wide",
)

# Palette CORRECTGLOSS : bleu nuit, turquoise et or.
st.markdown("""
<style>
:root { --cg-navy: #092B52; --cg-teal: #13B8A6; --cg-gold: #F6BE32; }
[data-testid="stAppViewContainer"] { background: #F6F9FC; }
[data-testid="stHeader"] { background: rgba(246,249,252,0.92); }
[data-testid="stSidebar"] { background: #EAF4F7; }
h1, h2, h3 { color: #092B52; }
[data-testid="stCaptionContainer"] { color: #456176; }
div.stButton > button[kind="primary"] { background: #13B8A6; border-color: #13B8A6; color: white; }
div.stButton > button[kind="primary"]:hover { background: #0E9F91; border-color: #0E9F91; color: white; }
[data-testid="stMetricValue"] { color: #092B52; }
</style>
""", unsafe_allow_html=True)

def render_brand_header(subtitle):
    logo_col, title_col = st.columns([1, 5])
    with logo_col:
        if LOGO_PATH.exists():
            st.image(str(LOGO_PATH), width=92)
        elif BRAND_PATH.exists():
            st.image(str(BRAND_PATH), width=92)
    with title_col:
        st.markdown(
            '<h1 style="margin-bottom:0; color:#092B52;">CORRECT<span style="color:#13B8A6;">GLOSS</span></h1>',
            unsafe_allow_html=True,
        )
        st.caption(subtitle)


# Protection facultative pour les déploiements web : définir APP_PASSWORD
# dans les variables d'environnement ou dans les secrets Streamlit.
def _configured_password():
    value = os.getenv("APP_PASSWORD", "").strip()
    if value:
        return value
    try:
        return str(st.secrets.get("APP_PASSWORD", "")).strip()
    except Exception:
        return ""

APP_PASSWORD = _configured_password()
if APP_PASSWORD and not st.session_state.get("authenticated", False):
    render_brand_header("Correcteur de mémoire universitaire")
    password_entered = st.text_input("Mot de passe d'accès", type="password")
    if st.button("Se connecter", type="primary"):
        if hmac.compare_digest(password_entered, APP_PASSWORD):
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Mot de passe incorrect.")
    st.stop()

render_brand_header("Correction • reformulation académique • mise en page • contrôle des citations")("Correction • reformulation académique • mise en page • contrôle des citations")
if not APP_PASSWORD:
    st.caption("Pour une mise en ligne, configurez APP_PASSWORD dans les secrets Streamlit avant de partager l’adresse.")
st.info("Prototype local/web. Conservez toujours une copie originale. Les vérifications bibliographiques sont des contrôles de cohérence, pas une preuve que chaque source existe.")

with st.sidebar:
    st.header("Paramètres")
    mode = st.radio("Mode de correction", ["Hors ligne", "En ligne (LanguageTool)", "En ligne (IA + LanguageTool)"])
    language = st.selectbox("Langue principale", ["Français", "Anglais", "Malagasy"], index=0)
    st.subheader("Mise en page")
    font = st.selectbox("Police", ["Times New Roman", "Arial", "Calibri"], index=0)
    font_size = st.number_input("Taille (pt)", 10, 14, 12)
    line_spacing = st.selectbox("Interligne", ["1.0", "1.15", "1.5", "2.0"], index=2)
    margin_cm = st.number_input("Marges (cm)", 1.5, 4.0, 2.5, 0.25)
    justify = st.checkbox("Justifier le texte", True)
    page_numbers_note = st.caption("Les numéros de page dépendent du logiciel qui ouvre le DOCX. Ajoutez-les dans Word/LibreOffice si nécessaire.")

st.markdown("### 1. Importer le mémoire")
uploaded = st.file_uploader("Choisissez un fichier Word (.docx)", type=["docx"])
if not uploaded:
    st.markdown("""
    **Fonctionnalités**
    - Contrôles locaux de base et repérage de formulations fréquentes.
    - Suggestions de correction LanguageTool en ligne.
    - Reformulation académique avec API IA facultative.
    - Mise en page DOCX.
    - Vérification de cohérence entre citations auteur-date et bibliographie.
    - Export DOCX et PDF (PDF via LibreOffice installé sur l'ordinateur).
    """)
    st.stop()

try:
    original_bytes = uploaded.getvalue()
    doc = Document(io.BytesIO(original_bytes))
except Exception as e:
    st.error(f"Impossible d'ouvrir ce document : {e}")
    st.stop()

def get_paragraphs(document):
    return [p.text for p in document.paragraphs]

def offline_suggestions(text):
    suggestions = []
    rules = [
        (r"\bmalgré que\b", "bien que / malgré le fait que", "Formulation généralement déconseillée dans un mémoire."),
        (r"\bau jour d'aujourd'hui\b", "aujourd'hui / à ce jour", "Redondance."),
        (r"\bce dernier\b", "Remplacer par le nom précis si le référent est ambigu.", "Référent potentiellement ambigu."),
        (r"\btrès important\b", "essentiel / majeur / déterminant (selon le sens)", "Style académique : privilégier un terme précis."),
        (r"\bon peut voir que\b", "les résultats montrent que / on observe que", "Style académique plus précis."),
        (r"\bje pense que\b", "les données suggèrent que / il semble que", "Éviter une opinion personnelle non justifiée."),
        (r"\bpar rapport à\b", "concernant / au regard de / en comparaison avec", "Vérifier le sens exact."),
        (r"\betc\.\b", "Préciser les éléments pertinents ou utiliser « etc. » avec parcimonie", "Éviter une liste imprécise."),
    ]
    for pattern, replacement, note in rules:
        for m in re.finditer(pattern, text, re.I):
            suggestions.append({"motif": m.group(0), "suggestion": replacement, "explication": note})
    # Heuristique prudente : détecte des espaces avant ponctuation courants, sans modifier le texte.
    for m in re.finditer(r"\s+[,.]", text):
        suggestions.append({"motif": m.group(0), "suggestion": m.group(0).strip(), "explication": "Espace superflu avant la ponctuation."})
    return suggestions

def languagetool_check(text, lang_code):
    url = "https://api.languagetool.org/v2/check"
    r = requests.post(url, data={"text": text[:20000], "language": lang_code}, timeout=25)
    r.raise_for_status()
    data = r.json()
    out = []
    for m in data.get("matches", []):
        replacements = [x.get("value", "") for x in m.get("replacements", [])[:4]]
        start, length = m.get("offset", 0), m.get("length", 0)
        out.append({
            "motif": text[start:start+length],
            "suggestion": " / ".join(replacements) if replacements else "Aucune suggestion automatique",
            "explication": m.get("message", "Suggestion LanguageTool"),
            "règle": m.get("rule", {}).get("id", "")
        })
    return out

def citations_and_bibliography(document):
    body = "\n".join(p.text for p in document.paragraphs)
    refs = []
    in_bib = False
    for p in document.paragraphs:
        if re.search(r"^(bibliographie|références bibliographiques|references|sources)\s*$", p.text.strip(), re.I):
            in_bib = True
            continue
        if in_bib and p.text.strip():
            refs.append(p.text.strip())
    # Captures approximatives : (Nom, 2020), Nom (2020), (Nom et Nom, 2020)
    cited = set()
    for m in re.finditer(r"\(([A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÿ'’\-]+(?:\s+(?:et|&)\s+[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÿ'’\-]+)?),?\s*((?:19|20)\d{2}[a-z]?)\)", body):
        cited.add((m.group(1).strip(), m.group(2)))
    for m in re.finditer(r"\b([A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÿ'’\-]+)\s*\(((?:19|20)\d{2}[a-z]?)\)", body):
        cited.add((m.group(1).strip(), m.group(2)))
    ref_text = "\n".join(refs)
    matched, missing = [], []
    for author, year in sorted(cited):
        if re.search(re.escape(author.split()[-1]), ref_text, re.I) and re.search(re.escape(year[:4]), ref_text):
            matched.append(f"{author} ({year})")
        else:
            missing.append(f"{author} ({year})")
    return refs, matched, missing

  """Ajoute un champ Table des matières Word actualisable."""
    title = document.add_paragraph()
    title.style = "Title"
    title.add_run("TABLE DES MATIÈRES")

    paragraph = document.add_paragraph()

    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")

    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = ' TOC \\o "1-3" \\h \\z \\u '

    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")

    text = OxmlElement("w:t")
    text.text = "Ouvrez dans Word et actualisez la table des matières."

    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")

    run._r.append(begin)
    run._r.append(instruction)
    run._r.append(separate)
    run._r.append(text)
    run._r.append(end)

    body = document._body._body
    body.remove(title._p)
    body.remove(paragraph._p)
    body.insert(0, paragraph._p)
    body.insert(0, title._p)

    settings = document.settings.element
    update_fields = settings.find(qn("w:updateFields"))

    if update_fields is None:
        update_fields = OxmlElement("w:updateFields")
        settings.append(update_fields)

    update_fields.set(qn("w:val"), "true")

    return documentdef format_document(document):
    for sec in document.sections:
        sec.top_margin = Inches(margin_cm / 2.54)
        sec.bottom_margin = Inches(margin_cm / 2.54)
        sec.left_margin = Inches(margin_cm / 2.54)
        sec.right_margin = Inches(margin_cm / 2.54)
    for p in document.paragraphs:
        if p.style and p.style.name.lower().startswith("heading"):
            pass
        else:
            if justify:
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        for run in p.runs:
            run.font.name = font
            run.font.size = Pt(int(font_size))
        p.paragraph_format.line_spacing = float(line_spacing)
        p.paragraph_format.space_after = Pt(6)
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.font.name = font
                        run.font.size = Pt(int(font_size))
    return document

paragraphs = get_paragraphs(doc)
nonempty = [p for p in paragraphs if p.strip()]
st.success(f"Document chargé : {len(paragraphs)} paragraphes, {len(nonempty)} paragraphes non vides.")
left, right = st.columns(2)
with left:
    st.subheader("Aperçu du texte")
    preview = "\n\n".join(nonempty[:8])
    st.text_area("Premiers paragraphes", preview, height=280)
with right:
    st.subheader("Contrôles")
    st.write(f"Nombre de mots approximatif : **{sum(len(p.split()) for p in paragraphs)}**")
    st.write(f"Nombre de tableaux : **{len(doc.tables)}**")
    st.write(f"Nombre de sections Word : **{len(doc.sections)}**")

st.markdown("---")
st.markdown("### 2. Correction et suggestions")
combined_text = "\n\n".join(nonempty)
if st.button("Analyser le mémoire", type="primary"):
    all_suggestions = []
    with st.spinner("Analyse en cours..."):
        all_suggestions.extend([{**x, "source": "Contrôle local"} for x in offline_suggestions(combined_text)])
        if mode in ["En ligne (LanguageTool)", "En ligne (IA + LanguageTool)"]:
            try:
                lang_code = {"Français": "fr", "Anglais": "en-US", "Malagasy": "auto"}[language]
                if lang_code != "auto":
                    all_suggestions.extend([{**x, "source": "LanguageTool"} for x in languagetool_check(combined_text, lang_code)])
                else:
                    st.warning("LanguageTool ne prend pas en charge le malagasy dans ce prototype. Seuls les contrôles locaux seront exécutés.")
            except Exception as e:
                st.warning(f"LanguageTool indisponible : {e}. Les contrôles locaux restent disponibles.")
    st.session_state["suggestions"] = all_suggestions
    st.session_state["refs_result"] = citations_and_bibliography(doc)
    st.session_state["analysis_done"] = True["analysis_done"] = True

if st.session_state.get("analysis_done"):
    suggestions = st.session_state.get("suggestions", [])
    st.write(f"Suggestions repérées : **{len(suggestions)}**")
    if suggestions:
        st.dataframe(suggestions, use_container_width=True, hide_index=True)
        st.caption("Les suggestions ne sont pas appliquées automatiquement : vérifiez le contexte et corrigez dans Word pour éviter de changer le sens.")
    else:
        st.info("Aucune suggestion détectée par les contrôles activés. Cela ne garantit pas l'absence d'erreurs.")
    refs, matched, missing = st.session_state.get("refs_result", ([], [], []))
    st.markdown("### 3. Cohérence citations / bibliographie")
    c1, c2, c3 = st.columns(3)
    c1.metric("Entrées bibliographiques repérées", len(refs))
    c2.metric("Citations appariées (approx.)", len(matched))
    c3.metric("Citations à vérifier", len(missing))
    if missing:
        st.warning("Ces citations ne semblent pas correspondre à une entrée bibliographique selon une recherche automatique approximative : " + ", ".join(missing))
    else:
        st.success("Aucune citation manquante détectée par ce contrôle approximatif.")
    if refs:
        with st.expander("Entrées de bibliographie détectées"):
            st.write("\n".join("- " + x for x in refs))
    st.caption("Le programme ne vérifie pas encore l'existence réelle des ouvrages, DOI, URL, ni la conformité complète à APA/ISO 690. Il compare seulement quelques motifs auteur-date.")

st.markdown("---")
st.markdown("### 4. Reformulation académique (facultative)")
st.write("Pour une reformulation IA, configurez une clé API dans les variables d'environnement. N'envoyez pas de données confidentielles sans autorisation.")
api_key = os.getenv("OPENAI_API_KEY", "")
rewrite_text = st.text_area("Collez ici un paragraphe à reformuler (facultatif)", height=150, placeholder="Collez un paragraphe de votre mémoire...")
if st.button("Reformuler ce paragraphe avec l'IA") and rewrite_text.strip():
    if not api_key:
        st.error("Clé OPENAI_API_KEY absente. Configurez-la dans l'environnement, puis relancez l'application.")
    else:
        try:
            from openai import OpenAI
            client = OpenAI(api_key=api_key)
            prompt = ("Reformule en français académique universitaire, sans inventer de faits, de chiffres ou de sources. "
                      "Conserve le sens et les termes techniques. Si le texte est ambigu, signale-le après la reformulation.\n\n" + rewrite_text)
            response = client.responses.create(model=os.getenv("OPENAI_MODEL", "gpt-4.1-mini"), input=prompt)
            st.text_area("Proposition de reformulation — à vérifier", response.output_text, height=220)
        except Exception as e:
            st.error(f"Échec de la reformulation : {e}")

st.markdown("---")
st.markdown("### 5. Export et mise en page")
st.write("Le fichier exporté est une copie formatée du document original. Les corrections de langue sont présentées sous forme de suggestions et ne sont pas injectées automatiquement.")
if st.button("Appliquer la mise en page et préparer le DOCX", type="primary"):
   formatted = Document(io.BytesIO(original_bytes))
formatted = format_document(formatted)
formatted = add_word_toc(formatted)

out = io.BytesIO()
    formatted.save(out)
    st.session_state["formatted_docx"] = out.getvalue()["formatted_docx"] = out.getvalue()
    st.success("Mise en page appliquée.")
if st.session_state.get("formatted_docx"):
    st.download_button("⬇️ Télécharger le mémoire mis en page (.docx)", st.session_state["formatted_docx"], file_name="memoire_mis_en_page.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

if st.session_state.get("formatted_docx"):
    if st.button("Préparer le PDF"):
        with st.spinner("Conversion PDF..."):
            soffice = shutil.which("soffice") or shutil.which("libreoffice")
            if not soffice:
                st.error("LibreOffice introuvable. Installez LibreOffice puis relancez l'application. Sinon ouvrez le DOCX dans Word et utilisez Fichier > Exporter > Créer PDF.")
            else:
                with tempfile.TemporaryDirectory() as td:
                    docx_path = Path(td) / "memoire_mis_en_page.docx"
                    docx_path.write_bytes(st.session_state["formatted_docx"])
                    try:
                        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", td, str(docx_path)], check=True, capture_output=True, timeout=90)
                        pdf_path = Path(td) / "memoire_mis_en_page.pdf"
                        if pdf_path.exists():
                            st.session_state["formatted_pdf"] = pdf_path.read_bytes()
                        else:
                            st.error("La conversion n'a pas produit de PDF. Vérifiez l'installation de LibreOffice.")
                    except Exception as e:
                        st.error(f"Erreur de conversion : {e}")
if st.session_state.get("formatted_pdf"):
    st.download_button("⬇️ Télécharger le mémoire mis en page (.pdf)", st.session_state["formatted_pdf"], file_name="memoire_mis_en_page.pdf", mime="application/pdf")

st.markdown("---")
st.caption("CORRECTGLOSS — version 0.3 • Traitez les résultats comme des suggestions, pas comme une validation académique.")

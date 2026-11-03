import io, os, re, tempfile, subprocess, shutil
from pathlib import Path
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import requests
from research_engine import search_web, search_bibliography
from correction_engine import offline_suggestions, languagetool_suggestions, apply_suggestions

APP_NAME = "Mémoire Correcteur Pro"
st.set_page_config(page_title=APP_NAME, page_icon="📘", layout="wide")

# Google Analytics : suivi des visites CORRECTGLOSS.
# Le code est injecté dans la page Streamlit (pas dans un iframe).
GA_MEASUREMENT_ID = "G-MPMT3MR3M6"
st.html(
    f"""
    <script async src="https://www.googletagmanager.com/gtag/js?id={GA_MEASUREMENT_ID}"></script>
    <script>
      window.dataLayer = window.dataLayer || [];
      function gtag(){{dataLayer.push(arguments);}}
      gtag('js', new Date());
      gtag('config', '{GA_MEASUREMENT_ID}');
    </script>
    """,
    unsafe_allow_javascript=True,
)

st.title("📘 CORRECTGLOSS")
st.caption("Correction • reformulation académique • mise en page • contrôle des citations")

st.markdown("""
<style>
/* Mise en page adaptative : conserve toutes les fonctions sur PC et smartphone. */
@media (max-width: 768px) {
  .main .block-container {
    padding-left: 0.85rem;
    padding-right: 0.85rem;
    padding-top: 1rem;
    padding-bottom: 2rem;
    max-width: 100%;
  }
  h1 { font-size: 1.65rem !important; }
  h2 { font-size: 1.35rem !important; }
  h3 { font-size: 1.15rem !important; }
  [data-testid="stButton"] button,
  [data-testid="stDownloadButton"] button {
    width: 100%;
    min-height: 2.8rem;
    white-space: normal;
  }
  [data-testid="stFileUploader"] section {
    padding: 0.75rem;
  }
  [data-testid="stDataFrame"],
  [data-testid="stTable"] {
    overflow-x: auto;
  }
  [data-testid="stTextArea"] textarea,
  [data-testid="stTextInput"] input {
    font-size: 16px !important;
  }
}
</style>
""", unsafe_allow_html=True)

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
            ["Ouvertures", "Recherches", "Analyses", "Exports Word/PDF"],
            ["app_open", "web_research", "document_analyzed", "docx_export_prepared"]):
            col.metric(label, counts.get(key, 0))
        if events:
            import pandas as pd
            df = pd.DataFrame(events)
            df["created_at"] = pd.to_datetime(df["created_at"], errors="coerce", utc=True)
            daily = df.dropna(subset=["created_at"]).assign(jour=lambda x: x["created_at"].dt.strftime("%Y-%m-%d"))
            daily = daily.groupby(["jour", "event_name"]).size().unstack(fill_value=0)
            st.subheader("Activité quotidienne")
            st.line_chart(daily)
            st.subheader("Répartition des fonctionnalités")
            st.bar_chart(df["event_name"].value_counts())
            st.subheader("Événements récents")
            st.dataframe(df, use_container_width=True, hide_index=True)
            st.caption("Les chiffres représentent des événements enregistrés, pas nécessairement des personnes uniques. Liste limitée aux 5 000 événements les plus récents.")
        else:
            st.info("Aucun événement enregistré pour le moment.")
    except Exception as e:
        st.error(f"Impossible de charger les statistiques : {e}")

def show_research_page():
    st.header("🌐 Recherche Internet & rédaction")
    st.write("Recherchez des informations, consultez les sources, puis préparez un brouillon à vérifier.")

    query = st.text_input(
        "Sujet de recherche",
        placeholder="Ex. : La compétitivité du girofle malgache à l'exportation",
        max_chars=300,
    )
    col1, col2 = st.columns([1, 2])
    with col1:
        max_results = st.selectbox("Nombre de sources", [5, 8, 10], index=1)
    with col2:
        research_language = st.selectbox("Langue des résultats", ["Français", "Anglais"], index=0)

    if st.button("🔎 Rechercher sur Internet", type="primary"):
        if not query.strip():
            st.warning("Saisissez d'abord un sujet.")
        else:
            search_query = query if research_language == "Français" else f"{query} English"
            try:
                with st.spinner("Recherche des sources en cours..."):
                    found = search_web(search_query, max_results=max_results)
                st.session_state["cg_research_query"] = query
                st.session_state["cg_research_results"] = found
                log_app_event("web_research")
            except Exception as exc:
                st.error(f"Recherche indisponible : {exc}")
                st.info("Réessayez plus tard ou vérifiez votre connexion Internet.")

    st.divider()
    st.subheader("📚 Auteurs, livres et documents académiques")
    st.write("Trouvez des références bibliographiques, des articles scientifiques, des livres, des mémoires et des thèses liés à votre thème.")
    if st.button("📖 Rechercher des auteurs, livres et documents"):
        if len(query.strip()) < 3:
            st.warning("Saisissez d'abord un thème de recherche.")
        else:
            try:
                with st.spinner("Recherche des références bibliographiques..."):
                    bibliography = search_bibliography(query, max_results=max_results)
                st.session_state["cg_bibliography"] = bibliography
                log_app_event("bibliography_search")
            except Exception as exc:
                st.error(f"Recherche bibliographique indisponible : {exc}")

    bibliography = st.session_state.get("cg_bibliography", {})
    if bibliography:
        st.info(bibliography.get("_message", "Résultats issus de catalogues publics."))
        diagnostics = bibliography.get("_diagnostics", [])
        if diagnostics:
            with st.expander("Détails des sources temporairement indisponibles"):
                for detail in diagnostics:
                    st.caption(detail)

    for category in ["Auteurs et chercheurs", "Livres", "Articles scientifiques", "Documents, mémoires et thèses"]:
        entries = bibliography.get(category, [])
        st.markdown(f"**{category} ({len(entries)})**")
        if not entries:
            st.caption("Aucun résultat pertinent trouvé. Essayez un thème plus précis ou une variante en anglais.")
        for item in entries:
            title = item.get("title") or item.get("name") or "Référence sans titre"
            url = item.get("url")
            st.markdown(f"- **[{title}]({url})**" if url else f"- **{title}**")
            if item.get("authors"):
                st.caption("Auteur(s) : " + ", ".join(item["authors"]))
            if item.get("year"):
                st.caption(f"Année : {item['year']}")
            if item.get("doi"):
                st.caption("DOI : " + str(item["doi"]).replace("https://doi.org/", ""))
            if item.get("relevance") is not None:
                st.caption(f"Pertinence estimée : {round(float(item['relevance']) * 100)} %")
            if item.get("snippet"):
                st.write(item["snippet"])
            st.caption("Catalogue : " + item.get("provider", "Catalogue en ligne"))

    if bibliography:
        st.warning("Les scores de pertinence sont indicatifs. Vérifiez la notice originale, l'édition, les auteurs, l'année et le contenu avant toute citation académique.")
        if st.button("📥 Préparer la bibliographie en Word et PDF"):
            try:
                bibliography_doc = Document()
                bibliography_doc.add_heading("Bibliographie – CORRECTGLOSS", 0)
                bibliography_doc.add_paragraph("Sujet : " + query)
                for category in ["Articles scientifiques", "Livres", "Documents, mémoires et thèses"]:
                    entries = bibliography.get(category, [])
                    if entries:
                        bibliography_doc.add_heading(category, level=1)
                        for item in entries:
                            p = bibliography_doc.add_paragraph(style="List Number")
                            p.add_run(item.get("title", "Référence sans titre")).bold = True
                            details = []
                            if item.get("authors"): details.append(", ".join(item["authors"]))
                            if item.get("year"): details.append(str(item["year"]))
                            if item.get("doi"): details.append("DOI : " + str(item["doi"]).replace("https://doi.org/", ""))
                            if item.get("url"): details.append(item["url"])
                            if details: bibliography_doc.add_paragraph(" — ".join(details))
                buffer = io.BytesIO()
                bibliography_doc.save(buffer)
                st.session_state["cg_bibliography_docx"] = buffer.getvalue()
                try:
                    st.session_state["cg_bibliography_pdf"] = convert_docx_to_pdf(buffer.getvalue())
                    st.session_state.pop("cg_bibliography_pdf_error", None)
                except Exception as pdf_error:
                    st.session_state.pop("cg_bibliography_pdf", None)
                    st.session_state["cg_bibliography_pdf_error"] = str(pdf_error)
                log_app_event("bibliography_export")
                st.success("Bibliographie Word préparée.")
            except Exception as export_error:
                st.error(f"Erreur pendant la préparation : {export_error}")
        if st.session_state.get("cg_bibliography_docx"):
            c1, c2 = st.columns(2)
            with c1:
                st.download_button("Télécharger bibliographie Word", st.session_state["cg_bibliography_docx"],
                    file_name="correctgloss_bibliographie.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            with c2:
                if st.session_state.get("cg_bibliography_pdf"):
                    st.download_button("Télécharger bibliographie PDF", st.session_state["cg_bibliography_pdf"],
                        file_name="correctgloss_bibliographie.pdf", mime="application/pdf")
            if st.session_state.get("cg_bibliography_pdf_error"):
                st.caption("Conversion PDF indisponible : " + st.session_state["cg_bibliography_pdf_error"])

    sources = st.session_state.get("cg_research_results", [])
    if sources:
        st.subheader(f"Sources trouvées : {len(sources)}")
        for index, source in enumerate(sources, start=1):
            st.markdown(f"**[{index}] [{source['title']}]({source['url']})**")
            if source["snippet"]:
                st.write(source["snippet"])
            st.caption(source["url"])

        st.warning("Les résultats et extraits sont des pistes. Ouvrez les liens, vérifiez l'auteur, la date et le contenu avant toute citation académique.")

        st.subheader("✍️ Préparer une rédaction")
        draft_type = st.selectbox(
            "Type de rédaction",
            ["Plan détaillé", "Introduction", "Synthèse documentaire", "Développement", "Conclusion", "Résumé"],
        )
        extra_instructions = st.text_area(
            "Consignes complémentaires (facultatif)",
            placeholder="Ex. : niveau Master, contexte de Madagascar, problématique...",
            height=90,
        )
        if st.button("Générer un brouillon avec les sources"):
            api_key = secret_value("OPENAI_API_KEY") or os.getenv("OPENAI_API_KEY", "")
            if not api_key:
                st.error("La clé OPENAI_API_KEY doit être configurée dans les secrets Streamlit.")
            else:
                try:
                    from openai import OpenAI
                    source_context = "\n\n".join(
                        f"[{i}] {item['title']}\nURL: {item['url']}\nExtrait: {item['snippet']}"
                        for i, item in enumerate(sources, start=1)
                    )
                    type_rules = {
                        "Plan détaillé": """PRODUIS UNIQUEMENT UN PLAN DÉTAILLÉ.
Présente une structure hiérarchisée avec PARTIES, CHAPITRES et sections numérotées.
Ajoute des intitulés précis et, si utile, une courte indication sous chaque grande partie.
N'écris PAS d'introduction rédigée, de développement, de conclusion, de résumé ni de bibliographie.""",
                        "Introduction": """PRODUIS UNIQUEMENT UNE INTRODUCTION ACADÉMIQUE RÉDIGÉE.
Structure-la avec le contexte et la justification du sujet, l'intérêt du sujet, la problématique,
les questions de recherche, les objectifs, les hypothèses si les informations le permettent,
la méthodologie annoncée et la présentation du plan en quelques phrases.
Ne rédige PAS le plan détaillé, les chapitres, le développement ni la conclusion.""",
                        "Synthèse documentaire": """PRODUIS UNIQUEMENT UNE SYNTHÈSE DOCUMENTAIRE.
Organise les idées et les connaissances issues des extraits, compare les points de vue
lorsque les sources le permettent et signale les limites. N'ajoute pas d'introduction générale,
de plan détaillé, de développement de mémoire ni de conclusion générale.""",
                        "Développement": """PRODUIS UNIQUEMENT LE DÉVELOPPEMENT du sujet, avec des titres et sous-titres pertinents.
Développe les idées soutenues par les extraits. N'écris ni l'introduction générale,
ni le plan complet, ni la conclusion générale.""",
                        "Conclusion": """PRODUIS UNIQUEMENT UNE CONCLUSION ACADÉMIQUE.
Fais une synthèse des éléments établis, réponds prudemment à la problématique si les extraits
le permettent, rappelle les limites et propose des perspectives. N'ajoute ni introduction,
ni plan détaillé, ni nouveau développement.""",
                        "Résumé": """PRODUIS UNIQUEMENT UN RÉSUMÉ fidèle et concis du sujet et des informations disponibles.
N'ajoute ni plan détaillé, ni introduction complète, ni développement chapitre par chapitre,
ni conclusion séparée."""
                    }
                    prompt = f"""Tu es un assistant de recherche et de rédaction académique.
Respecte STRICTEMENT le type de document demandé. Le type sélectionné détermine le seul contenu
à produire : ne le remplace pas par un autre type et ne mélange jamais plusieurs livrables.

SUJET : {st.session_state.get('cg_research_query', query)}
TYPE EXACT À PRODUIRE : {draft_type}

INSTRUCTIONS OBLIGATOIRES POUR CE TYPE :
{type_rules.get(draft_type, "Produis uniquement le type demandé.")}

Consignes complémentaires : {extra_instructions or 'Rédaction claire et structurée en français.'}

Utilise uniquement les informations présentes dans les extraits ci-dessous.
N'invente aucun fait, chiffre, auteur, date ou référence. Si les extraits ne suffisent pas,
indique brièvement les informations manquantes sans inventer. Place des appels de source [1], [2]
uniquement lorsque l'extrait correspondant soutient réellement l'information.
N'ajoute pas de section « Sources à vérifier » au milieu du livrable : termine par une courte
section « Sources à vérifier » après le contenu demandé, reprenant les URL fournies.
La section des sources ne doit pas transformer le livrable en un autre type de document.
Précise en une phrase finale qu'il s'agit d'un brouillon à vérifier avec les pages originales.

SOURCES :
{source_context}
"""
                    client = OpenAI(api_key=api_key)
                    with st.spinner("Préparation du brouillon..."):
                        response = client.responses.create(
                            model=os.getenv("OPENAI_MODEL", "gpt-4.1-mini"),
                            input=prompt,
                        )
                    st.session_state["cg_research_draft"] = response.output_text
                    log_app_event("research_draft_generated")
                except Exception as exc:
                    st.error(f"Impossible de générer le brouillon : {exc}")

    if st.session_state.get("cg_research_draft"):
        st.text_area("Brouillon à relire et vérifier", st.session_state["cg_research_draft"], height=420)
        st.download_button(
            "Télécharger le brouillon (.txt)",
            st.session_state["cg_research_draft"],
            file_name="correctgloss_brouillon_recherche.txt",
            mime="text/plain",
        )


def detect_and_style_headings(document):
    """Classe les titres en niveaux logiques sans modifier le texte du mémoire."""
    main_titles = {
        "INTRODUCTION", "CONCLUSION", "BIBLIOGRAPHIE", "SITOGRAPHIE",
        "ANNEXES", "REMERCIEMENTS", "RÉSUMÉ", "RESUME", "ABSTRACT",
        "TABLE DES MATIÈRES", "TABLE DES MATIERES"
    }
    heading1 = document.styles["Heading 1"]
    heading2 = document.styles["Heading 2"]
    heading3 = document.styles["Heading 3"]

    for para in document.paragraphs:
        value = re.sub(r"\s+", " ", para.text).strip()
        if not value:
            continue
        upper = value.upper()
        level = None

        if upper in main_titles or re.match(
            r"^(CHAPITRE|PARTIE)\s+([IVXLCDM]+|[0-9]+)\b", upper
        ):
            level = 1
        elif re.match(r"^\d+(?:\.\d+){2,}(?:\s|$)", value):
            level = 3
        elif re.match(r"^\d+\.\d+(?:\s|$)", value):
            level = 2
        elif re.match(r"^\d+\.(?:\s|$)", value):
            level = 1
        elif re.match(r"^(SECTION|SOUS[- ]PARTIE)\s+([IVXLCDM]+|[0-9]+)\b", upper):
            level = 2
        elif len(value) <= 100 and value == upper and len(value.split()) <= 12:
            # Titre court en majuscules : niveau principal, sauf s'il est déjà hiérarchisé.
            level = 1

        if level == 1:
            para.style = heading1
        elif level == 2:
            para.style = heading2
        elif level == 3:
            para.style = heading3

        if level:
            para.paragraph_format.keep_with_next = True
            para.paragraph_format.keep_together = True
            for run in para.runs:
                run.font.name = "Times New Roman"
                run.bold = True
                run.font.size = Pt(14 if level == 1 else 12)

    return document


def add_word_toc(document):
    """Ajoute un champ TOC Word et demande sa mise à jour automatique à l'ouverture."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    settings = document.settings.element
    update_fields = settings.find(qn("w:updateFields"))
    if update_fields is None:
        update_fields = OxmlElement("w:updateFields")
        settings.append(update_fields)
    update_fields.set(qn("w:val"), "true")

    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title.add_run("TABLE DES MATIÈRES")
    title_run.bold = True
    title_run.font.name = "Times New Roman"
    title_run.font.size = Pt(16)

    p = document.add_paragraph()
    run = p.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    begin.set(qn("w:dirty"), "true")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = r' TOC \o "1-3" \h \z \u '
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    placeholder = OxmlElement("w:r")
    placeholder_text = OxmlElement("w:t")
    placeholder_text.text = "La table des matières sera générée automatiquement à l'ouverture dans Word."
    placeholder.append(placeholder_text)
    end_run = OxmlElement("w:r")
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    end_run.append(end)

    run._r.append(begin)
    run._r.append(instr)
    run._r.append(separate)
    p._p.append(placeholder)
    p._p.append(end_run)

    body = document._element.body
    body.remove(title._p)
    body.remove(p._p)
    # Place la table des matières à la toute fin du document,
    # juste avant les propriétés de section Word (sectPr).
    section_properties = body.sectPr
    if section_properties is not None:
        insert_at = list(body).index(section_properties)
        body.insert(insert_at, title._p)
        body.insert(insert_at + 1, p._p)
    else:
        body.append(title._p)
        body.append(p._p)
    return document

def refresh_docx_toc_with_libreoffice(docx_bytes):
    """Ouvre le DOCX avec LibreOffice UNO, actualise les index, puis l'enregistre."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    system_python = shutil.which("python3")
    if not soffice or not system_python:
        raise RuntimeError("LibreOffice et Python système sont nécessaires pour finaliser la table des matières.")

    with tempfile.TemporaryDirectory() as workdir:
        root = Path(workdir)
        source_file = root / "correctgloss_memoire.docx"
        output_file = root / "correctgloss_memoire_final.docx"
        source_file.write_bytes(docx_bytes)
        port = 2002
        profile_dir = root / "lo-profile"

        uno_script = root / "refresh_toc.py"
        uno_script.write_text(r'''
import sys, time
import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.connection import NoConnectException

def prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p

local_ctx = uno.getComponentContext()
resolver = local_ctx.ServiceManager.createInstanceWithContext(
    "com.sun.star.bridge.UnoUrlResolver", local_ctx)
ctx = None
for attempt in range(40):
    try:
        ctx = resolver.resolve(
            "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        break
    except NoConnectException:
        time.sleep(0.5)
if ctx is None:
    raise RuntimeError("Connexion à LibreOffice impossible")

smgr = ctx.ServiceManager
desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
source = uno.systemPathToFileUrl(sys.argv[1])
target = uno.systemPathToFileUrl(sys.argv[2])
doc = desktop.loadComponentFromURL(source, "_blank", 0, (prop("Hidden", True),))
if doc is None:
    raise RuntimeError("Impossible d'ouvrir le document Word")
try:
    indexes = doc.getDocumentIndexes()
    for i in range(indexes.getCount()):
        indexes.getByIndex(i).update()
    doc.storeAsURL(target, (prop("FilterName", "Office Open XML Text"),))
finally:
    doc.close(True)
''', encoding="utf-8")

        launch = subprocess.Popen([
            soffice, "--headless",
            f"-env:UserInstallation={profile_dir.as_uri()}",
            f"--accept=socket,host=localhost,port={port};urp;StarOffice.ServiceManager",
            "--norestore", "--nodefault", "--nofirststartwizard"
        ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        try:
            env = os.environ.copy()
            env["PYTHONPATH"] = "/usr/lib/python3/dist-packages"
            result = subprocess.run(
                [system_python, str(uno_script), str(source_file), str(output_file)],
                capture_output=True, text=True, timeout=150, env=env
            )
            if result.returncode != 0 or not output_file.exists():
                detail = (result.stderr or result.stdout or "Erreur UNO inconnue")[-1500:]
                raise RuntimeError(detail)
            return output_file.read_bytes()
        finally:
            launch.terminate()
            try:
                launch.wait(timeout=5)
            except subprocess.TimeoutExpired:
                launch.kill()


def convert_docx_to_pdf(docx_bytes):
    """Convertit un DOCX finalisé en PDF pour une ouverture facile sur PC et Android."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise RuntimeError("LibreOffice est nécessaire pour créer le PDF.")
    with tempfile.TemporaryDirectory() as workdir:
        root = Path(workdir)
        source_file = root / "correctgloss_memoire.docx"
        output_file = root / "correctgloss_memoire.pdf"
        profile_dir = root / "lo-pdf-profile"
        source_file.write_bytes(docx_bytes)
        result = subprocess.run([
            soffice, "--headless",
            f"-env:UserInstallation={profile_dir.as_uri()}",
            "--convert-to", "pdf", "--outdir", str(root), str(source_file),
        ], capture_output=True, text=True, timeout=150)
        if result.returncode != 0 or not output_file.exists():
            detail = (result.stderr or result.stdout or "Erreur de conversion PDF inconnue")[-1200:]
            raise RuntimeError(detail)
        return output_file.read_bytes()

with st.sidebar:
    page = st.radio("Navigation", ["Correction", "Recherche & rédaction", "Dashboard administrateur"])
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

if page == "Recherche & rédaction":
    show_research_page()
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

if st.button("Analyser le mémoire", type="primary"):
    results = offline_suggestions(text, language=language)
    if mode == "En ligne (LanguageTool)" and language != "Malagasy":
        code = "fr" if language == "Français" else "en-US"
        try:
            with st.spinner("Vérification grammaticale en ligne..."):
                results.extend(languagetool_suggestions(text, code))
        except Exception as e:
            st.warning(f"LanguageTool indisponible : {e}")
    elif language == "Malagasy" and mode != "Hors ligne":
        st.warning("LanguageTool n'est pas activé pour le malagasy. Les règles locales restent disponibles.")
    # Identifiant stable pour conserver les choix de l'utilisateur lors du rerun Streamlit.
    for i, item in enumerate(results):
        item["id"] = i
    st.session_state["cg_results"] = results
    st.session_state["cg_corrected_text"] = text
    st.session_state["cg_original_text"] = text
    log_app_event("document_analyzed")

if "cg_results" in st.session_state:
    st.subheader("Suggestions de correction")
    results = st.session_state["cg_results"]
    st.write(f"{len(results)} suggestion(s) détectée(s). Le texte original est conservé.")
    if results:
        accepted = []
        for item in results:
            label = f"{item['category']} · « {item['motif']} » → « {item['suggestion']} »"
            checked = st.checkbox(label, key=f"cg_accept_{item['id']}")
            st.caption(item.get("explication", "Suggestion à vérifier."))
            if checked:
                accepted.append(item["id"])
        col_apply, col_all, col_reset = st.columns(3)
        with col_apply:
            if st.button("Appliquer les corrections sélectionnées", type="primary"):
                chosen = [x for x in results if x["id"] in accepted]
                st.session_state["cg_corrected_text"] = apply_suggestions(
                    st.session_state["cg_corrected_text"], chosen
                )
                st.success(f"{len(chosen)} correction(s) appliquée(s).")
        with col_all:
            if st.button("Tout corriger automatiquement"):
                st.session_state["cg_corrected_text"] = apply_suggestions(
                    st.session_state["cg_corrected_text"], results
                )
                st.success("Toutes les suggestions applicables ont été intégrées.")
        with col_reset:
            if st.button("Réinitialiser au texte original"):
                st.session_state["cg_corrected_text"] = st.session_state["cg_original_text"]
                st.success("Texte original restauré.")
    else:
        st.info("Aucune suggestion automatique trouvée. Cela ne garantit pas l'absence d'erreurs.")
    st.text_area("Texte après correction", st.session_state.get("cg_corrected_text", text), height=300)

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
    try:
        ready_docx = refresh_docx_toc_with_libreoffice(output.getvalue())
        st.session_state["cg_docx"] = ready_docx
        st.session_state.pop("cg_pdf", None)
        st.session_state.pop("cg_pdf_error", None)
        try:
            st.session_state["cg_pdf"] = convert_docx_to_pdf(ready_docx)
            log_app_event("pdf_export_prepared")
        except Exception as pdf_exc:
            st.session_state["cg_pdf_error"] = str(pdf_exc)
        log_app_event("docx_export_prepared")
        st.success("DOCX prêt. Vous pouvez aussi télécharger une version PDF compatible PC et Android.")
    except Exception as exc:
        st.error(f"La table des matières n'a pas pu être actualisée automatiquement : {exc}")
        st.info("Le fichier n'est pas proposé au téléchargement afin d'éviter une table incomplète.")

if "cg_docx" in st.session_state:
    col_docx, col_pdf = st.columns(2)
    with col_docx:
        st.download_button("Télécharger le Word (.docx)", st.session_state["cg_docx"],
            file_name="correctgloss_memoire.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True)
    with col_pdf:
        if st.session_state.get("cg_pdf"):
            st.download_button("Télécharger le PDF (.pdf)", st.session_state["cg_pdf"],
                file_name="correctgloss_memoire.pdf", mime="application/pdf",
                use_container_width=True)
        else:
            st.button("PDF indisponible", disabled=True, use_container_width=True)
    if st.session_state.get("cg_pdf_error"):
        st.warning("Le Word est prêt, mais la conversion PDF a échoué : " + st.session_state["cg_pdf_error"])
    st.caption("PDF : recommandé pour une ouverture facile sur Android et PC. Word : recommandé pour modifier le document.")

st.caption("Les suggestions sont indicatives. Conservez toujours une copie originale du mémoire.")

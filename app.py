import io
import os
import re
from datetime import datetime, timezone
from pathlib import Path

import requests
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from xml.sax.saxutils import escape as xml_escape
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT, TA_JUSTIFY
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

APP_NAME = "CORRECTGLOSS"
APP_VERSION = "0.5"
st.set_page_config(page_title=APP_NAME, page_icon="📘", layout="wide")

# -------------------------------------------------------------------
# CONFIGURATION
# -------------------------------------------------------------------
def get_secret(name: str, default: str = "") -> str:
    """Read a Streamlit secret first, then an environment variable."""
    try:
        value = st.secrets.get(name, default)
        if value:
            return str(value)
    except Exception:
        pass
    return os.getenv(name, default)


def log_app_event(event_type: str) -> None:
    """Log an event. The anon key must have a restricted INSERT-only RLS policy."""
    url = get_secret("SUPABASE_URL").rstrip("/")
    anon = get_secret("SUPABASE_ANON_KEY")
    if not url or not anon:
        return
    try:
        requests.post(
            f"{url}/rest/v1/app_events",
            headers={
                "apikey": anon,
                "Authorization": f"Bearer {anon}",
                "Content-Type": "application/json",
                "Prefer": "return=minimal",
            },
            json={"event_name": event_type, "app_version": APP_VERSION},
            timeout=5,
        ).raise_for_status()
    except Exception:
        # Logging must never prevent the app from working.
        pass


if "cg_session_logged" not in st.session_state:
    log_app_event("app_open")
    st.session_state["cg_session_logged"] = True


# -------------------------------------------------------------------
# HELPERS
# -------------------------------------------------------------------
def offline_suggestions(content: str):
    rules = [
        (r"\bmalgré que\b", "bien que / malgré le fait que", "Formulation à vérifier."),
        (r"\bau jour d'aujourd'hui\b", "aujourd'hui / à ce jour", "Redondance."),
        (r"\btrès important\b", "essentiel / majeur / déterminant", "Préciser le sens."),
        (r"\bon peut voir que\b", "les résultats montrent que", "Style académique."),
        (r"\bje pense que\b", "les données suggèrent que", "Éviter une opinion non justifiée."),
    ]
    found = []
    for pattern, suggestion, note in rules:
        for match in re.finditer(pattern, content, re.I):
            found.append({
                "motif": match.group(),
                "suggestion": suggestion,
                "explication": note,
            })
    return found



def generate_free_research(task: str, subject: str, context: str, language: str) -> str:
    """Build academic writing templates locally, without a paid API."""
    subject = subject.strip()
    context = context.strip() or "[À compléter avec les informations vérifiées de votre terrain]"
    if language == "Français":
        if task == "Proposer un plan détaillé":
            return f"""PLAN DÉTAILLÉ — PROPOSITION À ADAPTER
Sujet : {subject}

INTRODUCTION GÉNÉRALE
- Contexte et justification du sujet
- Problématique et question principale
- Objectif général et objectifs spécifiques
- Hypothèses de recherche
- Méthodologie et délimitation du terrain
- Annonce du plan

PARTIE I — CADRE CONCEPTUEL ET THÉORIQUE
Chapitre 1 : Définition des concepts clés
Chapitre 2 : Théories et travaux antérieurs

PARTIE II — CADRE MÉTHODOLOGIQUE ET ANALYSE
Chapitre 3 : Présentation du terrain et méthodologie
Chapitre 4 : Présentation et analyse des données collectées

PARTIE III — DISCUSSION ET RECOMMANDATIONS
Chapitre 5 : Interprétation et discussion des résultats
Chapitre 6 : Recommandations, limites et perspectives

CONCLUSION GÉNÉRALE
Synthèse, réponse à la problématique et perspectives.

Contexte à intégrer : {context}
Adaptez ce plan aux exigences de votre établissement et à vos données réelles."""
        if task == "Rédiger une introduction":
            return f"""INTRODUCTION — TRAME À COMPLÉTER
Sujet : {subject}

1. Contexte général : présentez le domaine avec des faits vérifiés et des sources datées.
2. Délimitation : précisez le secteur, l'organisation, la période et la zone étudiés.
3. Constat : décrivez le problème observé et justifiez le choix du sujet.
4. Problématique : expliquez précisément la difficulté ou le manque de connaissance.
5. Question principale : quels facteurs influencent {subject.lower()} dans le contexte étudié ?
6. Objectif général : analyser {subject.lower()}.
7. Objectifs spécifiques : décrire la situation, identifier les facteurs, analyser les résultats et proposer des améliorations.
8. Méthodologie : précisez les méthodes de collecte et d'analyse des données.
9. Annonce du plan : présentez brièvement les grandes parties.

Contexte à intégrer : {context}
Remplacez chaque indication par des informations vérifiées."""
        if task == "Formuler une problématique et des hypothèses":
            return f"""PROBLÉMATIQUE ET HYPOTHÈSES — PROPOSITION
Sujet : {subject}

Constat de départ : [Décrire la situation observable avec des éléments vérifiables.]
Problème de recherche : [Expliquer la difficulté ou le manque de connaissance à étudier.]

Question principale :
Quels sont les principaux facteurs qui influencent {subject.lower()} dans le terrain retenu ?

Questions spécifiques :
1. Quelle est la situation actuelle ?
2. Quels facteurs économiques, organisationnels, techniques ou humains interviennent ?
3. Quelles conséquences ou contraintes peut-on constater ?
4. Quelles améliorations peuvent être proposées ?

Hypothèses provisoires à vérifier :
H1 : Les pratiques organisationnelles et opérationnelles influencent la situation étudiée.
H2 : L'amélioration des méthodes et des ressources pourrait contribuer à de meilleurs résultats.

Contexte : {context}
Rendez les hypothèses précises, mesurables et vérifiables à partir de vos données."""
        if task == "Améliorer un paragraphe académique":
            return f"""AIDE À LA RÉVISION ACADÉMIQUE
Sujet : {subject}

Texte / contexte fourni :
{context}

Grille de révision :
1. Formulez une seule idée principale par paragraphe.
2. Développez cette idée avec un argument clair.
3. Ajoutez une preuve, un exemple ou une source vérifiable.
4. Expliquez le lien entre la preuve et la problématique.
5. Terminez par une transition logique.

Évitez les répétitions, les affirmations sans preuve et les formulations vagues.
Cette version gratuite fournit une méthode de révision, pas une réécriture automatique par IA."""
        if task == "Préparer une conclusion":
            return f"""CONCLUSION — TRAME À COMPLÉTER
Sujet : {subject}

1. Rappelez l'objectif et la démarche méthodologique.
2. Synthétisez les principaux résultats réellement obtenus.
3. Répondez directement à la question principale.
4. Indiquez si chaque hypothèse est confirmée, partiellement confirmée ou rejetée.
5. Formulez des recommandations liées aux résultats.
6. Présentez les limites et les perspectives de recherche.

Résultats et contexte à intégrer : {context}
N'ajoutez pas de nouvelles données non analysées dans le mémoire."""
        return f"""QUESTIONS DE RECHERCHE
Sujet : {subject}

Question principale :
Quels sont les facteurs, les effets et les pistes d'amélioration liés à {subject.lower()} ?

Questions spécifiques :
1. Quelle est la situation actuelle sur le terrain ?
2. Quels sont les acteurs concernés et leurs rôles ?
3. Quels facteurs expliquent la situation observée ?
4. Quelles contraintes et opportunités ressortent de l'étude ?
5. Quelles actions pourraient améliorer la situation ?

Contexte : {context}
Rendez chaque question précise, réalisable et cohérente avec votre méthodologie."""
    # For other selected languages, return a useful outline and clearly state the language limitation.
    return f"""FREE OFFLINE ACADEMIC TEMPLATE
Topic: {subject}
Selected language: {language}

Task: {task}

Suggested structure:
1. Background and context
2. Research problem and main question
3. General and specific objectives
4. Working hypotheses (if required)
5. Methodology and study area
6. Analysis of verified findings
7. Recommendations and limitations
8. Conclusion

Context to use:
{context}

Note: This free offline template is currently provided in French/English headings only. Adapt the wording into {language} and verify all facts, data and references before using it."""



def search_academic_sources(query: str, limit: int = 6):
    """Search Crossref and OpenAlex for real scholarly metadata; never invent references."""
    found = []
    seen = set()
    try:
        response = requests.get(
            "https://api.crossref.org/works",
            params={"query": query, "rows": limit, "select": "title,author,published,created,container-title,DOI,type,URL"},
            headers={"User-Agent": "CORRECTGLOSS/0.6 (academic source discovery)"},
            timeout=18,
        )
        response.raise_for_status()
        items = response.json().get("message", {}).get("items", [])
        for item in items:
            title = (item.get("title") or [""])[0].strip()
            if not title:
                continue
            doi = item.get("DOI", "")
            url = item.get("URL") or (f"https://doi.org/{doi}" if doi else "")
            authors = []
            for author in item.get("author", [])[:5]:
                name = " ".join(x for x in [author.get("given", ""), author.get("family", "")] if x)
                if name:
                    authors.append(name)
            date_parts = (item.get("published", {}) or {}).get("date-parts", [[]])[0]
            year = str(date_parts[0]) if date_parts else ""
            key = doi.lower() or title.lower()
            if key not in seen:
                seen.add(key)
                found.append({"title": title, "authors": ", ".join(authors) or "Auteur non indiqué",
                              "year": year, "publisher": (item.get("container-title") or [""])[0],
                              "doi": doi, "url": url, "source": "Crossref"})
    except Exception:
        pass
    try:
        response = requests.get(
            "https://api.openalex.org/works",
            params={"search": query, "per-page": limit, "select": "title,authorships,publication_year,doi,primary_location"},
            timeout=18,
        )
        response.raise_for_status()
        items = response.json().get("results", [])
        for item in items:
            title = (item.get("title") or "").strip()
            if not title:
                continue
            doi = (item.get("doi") or "").replace("https://doi.org/", "")
            url = item.get("doi") or ((item.get("primary_location") or {}).get("landing_page_url") or "")
            authors = []
            for authorship in item.get("authorships", [])[:5]:
                author = (authorship.get("author") or {}).get("display_name")
                if author:
                    authors.append(author)
            key = doi.lower() or title.lower()
            if key not in seen:
                seen.add(key)
                found.append({"title": title, "authors": ", ".join(authors) or "Auteur non indiqué",
                              "year": str(item.get("publication_year") or ""),
                              "publisher": "", "doi": doi, "url": url, "source": "OpenAlex"})
    except Exception:
        pass
    return found[:limit * 2]


def make_research_docx(title: str, body: str, sources: list):
    """Create a clean Word document from generated research text and verified metadata."""
    doc = Document()
    doc.add_heading(title or "Travail académique", 0)
    for block in re.split(r"\n\s*\n", body.strip()):
        block = block.strip()
        if not block:
            continue
        lines = block.splitlines()
        heading = lines[0].strip()
        if len(lines) == 1 and (heading.isupper() or heading.startswith(("PARTIE ", "CHAPITRE "))):
            doc.add_heading(heading, level=1)
        else:
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                if line.isupper() and len(line) < 100:
                    doc.add_heading(line, level=2)
                else:
                    doc.add_paragraph(line)
    if sources:
        doc.add_page_break()
        doc.add_heading("Bibliographie et sitographie — sources repérées", level=1)
        for source in sources:
            citation = f"{source.get('authors','Auteur non indiqué')} ({source.get('year','s.d.')}). {source.get('title','')}. {source.get('publisher','')}".strip()
            if source.get("doi"):
                citation += f" DOI : https://doi.org/{source['doi']}"
            elif source.get("url"):
                citation += f" URL : {source['url']}"
            doc.add_paragraph(citation)
    return doc

def offline_rewrite(text: str) -> str:
    """Apply a few transparent local text cleanups; no AI service required."""
    cleaned = re.sub(r"[ \\t]+", " ", text.strip())
    cleaned = re.sub(r"\\n{3,}", "\\n\\n", cleaned)
    for pattern, replacement in [
        (r"\\bau jour d'aujourd'hui\\b", "aujourd'hui"),
        (r"\\bmalgré que\\b", "bien que"),
        (r"\\bon peut voir que\\b", "les résultats montrent que"),
    ]:
        cleaned = re.sub(pattern, replacement, cleaned, flags=re.I)
    return cleaned


def call_ai(prompt: str) -> str:
    key = get_secret("OPENAI_API_KEY")
    if not key:
        raise RuntimeError("OPENAI_API_KEY n'est pas configurée dans les Secrets.")
    from openai import OpenAI
    client = OpenAI(api_key=key)
    model = get_secret("OPENAI_MODEL", "gpt-4.1-mini")
    response = client.responses.create(model=model, input=prompt)
    return response.output_text


def add_word_toc(document):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    paragraph = document.add_paragraph()
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = ' TOC \\o "1-3" \\h \\z \\u '
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text_run = OxmlElement("w:r")
    text_node = OxmlElement("w:t")
    text_node.text = "Ouvrez dans Word puis mettez à jour la table."
    text_run.append(text_node)
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.append(begin)
    run._r.append(instruction)
    run._r.append(separate)
    paragraph._p.append(text_run)
    run._r.append(end)
    body = document._element.body
    body.remove(paragraph._p)
    body.insert(0, paragraph._p)
    return document


def document_to_pdf_bytes(document, font_name: str, font_size: int, line_spacing: str, margin_cm: float, justify: bool) -> bytes:
    """Convert a python-docx document to a simple, readable A4 PDF."""
    output = io.BytesIO()
    margin = margin_cm * cm
    pdf = SimpleDocTemplate(
        output, pagesize=A4, rightMargin=margin, leftMargin=margin,
        topMargin=margin, bottomMargin=margin, title="Document CORRECTGLOSS",
    )
    base = getSampleStyleSheet()
    pdf_font = "Times-Roman" if font_name == "Times New Roman" else "Helvetica"
    leading = int(font_size * float(line_spacing))
    body_style = ParagraphStyle(
        "CGBody", parent=base["BodyText"], fontName=pdf_font,
        fontSize=font_size, leading=leading,
        alignment=TA_JUSTIFY if justify else TA_LEFT, spaceAfter=6,
    )
    heading_styles = {
        level: ParagraphStyle(
            f"CGHeading{level}", parent=base["Heading1"], fontName=pdf_font,
            fontSize=max(font_size, 16 - level), leading=max(font_size + 3, 19 - level),
            spaceBefore=10, spaceAfter=6, keepWithNext=True,
        )
        for level in range(1, 4)
    }
    story = []
    for paragraph in document.paragraphs:
        raw = paragraph.text.strip()
        if not raw:
            story.append(Spacer(1, leading * 0.35))
            continue
        style_name = (paragraph.style.name if paragraph.style else "").lower()
        match = re.search(r"(\d+)", style_name) if style_name.startswith("heading") else None
        level = min(int(match.group(1)), 3) if match else (1 if style_name.startswith("heading") else 0)
        safe_text = xml_escape(raw).replace("\n", "<br/>")
        story.append(Paragraph(safe_text, heading_styles[level] if level else body_style))
    for table in document.tables:
        rows = [
            [Paragraph(xml_escape(cell.text).replace("\n", "<br/>"), body_style) for cell in row.cells]
            for row in table.rows
        ]
        if rows:
            pdf_table = Table(rows, repeatRows=1, hAlign="LEFT")
            pdf_table.setStyle(TableStyle([
                ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]))
            story.extend([pdf_table, Spacer(1, 8)])
    pdf.build(story)
    return output.getvalue()


def admin_dashboard():
    st.subheader("Dashboard administrateur")
    expected_password = get_secret("ADMIN_PASSWORD")
    if not expected_password:
        st.warning(
            "Le dashboard est verrouillé : ajoutez ADMIN_PASSWORD dans "
            "Streamlit Cloud → Settings → Secrets."
        )
        return

    if not st.session_state.get("cg_admin_authenticated", False):
        password = st.text_input("Mot de passe administrateur", type="password")
        if st.button("Se connecter au dashboard", type="primary"):
            if password and password == expected_password:
                st.session_state["cg_admin_authenticated"] = True
                st.rerun()
            else:
                st.error("Mot de passe incorrect.")
        return

    if st.button("Se déconnecter du dashboard"):
        st.session_state["cg_admin_authenticated"] = False
        st.rerun()

    service_key = get_secret("SUPABASE_SERVICE_ROLE_KEY")
    url = get_secret("SUPABASE_URL").rstrip("/")
    if not url or not service_key:
        st.info(
            "Connexion Supabase administrateur non configurée. "
            "Ajoutez SUPABASE_URL et SUPABASE_SERVICE_ROLE_KEY dans les Secrets "
            "pour afficher les statistiques."
        )
        st.caption(
            "Ne partagez jamais la clé service_role et ne l'insérez jamais "
            "dans le code source public."
        )
        return

    try:
        response = requests.get(
            f"{url}/rest/v1/app_events",
            headers={
                "apikey": service_key,
                "Authorization": f"Bearer {service_key}",
            },
            params={
                "select": "event_name,app_version,created_at",
                "order": "created_at.desc",
                "limit": "5000",
            },
            timeout=15,
        )
        response.raise_for_status()
        rows = response.json()
        total = len(rows)
        opens = sum(1 for row in rows if row.get("event_name") == "app_open")
        analyses = sum(1 for row in rows if row.get("event_name") == "document_analyzed")
        rewrites = sum(1 for row in rows if row.get("event_name") == "ai_rewrite")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Événements chargés", total)
        c2.metric("Ouvertures", opens)
        c3.metric("Analyses", analyses)
        c4.metric("Reformulations IA", rewrites)
        st.caption("Statistiques basées sur les 5 000 événements les plus récents.")
        if rows:
            st.dataframe(rows, use_container_width=True, hide_index=True)
        else:
            st.info("Aucun événement enregistré pour le moment.")
    except Exception as exc:
        st.error(f"Impossible de lire les statistiques Supabase : {exc}")
        st.caption(
            "Vérifiez que la table app_events contient event_name, app_version "
            "et created_at, et que la clé service_role est correcte."
        )


# -------------------------------------------------------------------
# SIDEBAR NAVIGATION
# -------------------------------------------------------------------
with st.sidebar:
    st.title("📘 CORRECTGLOSS")
    st.caption("Assistant de rédaction académique")
    page = st.radio(
        "Navigation",
        [
            "Accueil",
            "Correction",
            "Recherche & rédaction",
            "Export",
            "Dashboard administrateur",
        ],
        index=0,
    )
    st.divider()
    st.caption(f"Version {APP_VERSION}")


# -------------------------------------------------------------------
# PAGE: HOME
# -------------------------------------------------------------------
st.title("📘 CORRECTGLOSS")
st.caption("Correction • reformulation académique • recherche et rédaction • export")

if page == "Accueil":
    st.header("Bienvenue sur CORRECTGLOSS")
    st.write(
        "CORRECTGLOSS vous accompagne dans la préparation de vos mémoires, "
        "rapports de stage et travaux académiques."
    )
    st.markdown(
        """
        **Fonctionnalités disponibles**
        - Correction linguistique hors ligne et avec LanguageTool
        - Reformulation en français académique avec IA
        - Aide à la recherche et à la rédaction
        - Mise en page et export Word
        - Espace administrateur protégé
        """
    )
    a, b, c = st.columns(3)
    with a:
        st.metric("Version", APP_VERSION)
    with b:
        st.metric("Langues de contrôle", "Français / Anglais / Malagasy*")
    with c:
        st.metric("Export disponible", "Word (.docx)")
    st.caption("*Le contrôle automatique en malagasy reste limité dans cette version.")
    if st.button("Commencer une correction", type="primary"):
        st.session_state["cg_page_choice"] = "Correction"
        st.info("Choisissez « Correction » dans le menu Navigation à gauche.")


# -------------------------------------------------------------------
# PAGE: CORRECTION
# -------------------------------------------------------------------
elif page == "Correction":
    st.header("Correction de document")
    with st.expander("Paramètres de correction", expanded=True):
        mode = st.radio(
            "Mode de correction",
            ["Hors ligne", "En ligne (LanguageTool)"],
            horizontal=True,
        )
        language = st.selectbox("Langue", ["Français", "Anglais", "Malagasy"])
    uploaded = st.file_uploader(
        "Importer un mémoire Word (.docx)",
        type=["docx"],
        key="correction_upload",
    )

    if uploaded:
        try:
            original_bytes = uploaded.getvalue()
            doc = Document(io.BytesIO(original_bytes))
            paragraphs = [p.text for p in doc.paragraphs]
            document_text = "\n\n".join(p for p in paragraphs if p.strip())
            st.success(
                f"Document chargé : {len(paragraphs)} paragraphes • "
                f"{len(doc.tables)} tableaux"
            )
            st.text_area(
                "Aperçu du document",
                "\n\n".join([p for p in paragraphs if p.strip()][:8]),
                height=220,
                disabled=True,
            )

            if st.button("Analyser le mémoire", type="primary"):
                results = offline_suggestions(document_text)
                if mode == "En ligne (LanguageTool)" and language != "Malagasy":
                    code = "fr" if language == "Français" else "en-US"
                    try:
                        response = requests.post(
                            "https://api.languagetool.org/v2/check",
                            data={"text": document_text[:20000], "language": code},
                            timeout=25,
                        )
                        response.raise_for_status()
                        for item in response.json().get("matches", []):
                            start = item.get("offset", 0)
                            length = item.get("length", 0)
                            replacements = [
                                item2.get("value", "")
                                for item2 in item.get("replacements", [])[:3]
                            ]
                            results.append({
                                "motif": document_text[start:start + length],
                                "suggestion": " / ".join(replacements) or "À vérifier",
                                "explication": item.get("message", "LanguageTool"),
                            })
                    except Exception as exc:
                        st.warning(f"LanguageTool indisponible : {exc}")
                elif language == "Malagasy" and mode != "Hors ligne":
                    st.warning(
                        "LanguageTool n'est pas activé pour le malagasy "
                        "dans cette version."
                    )
                st.session_state["cg_results"] = results
                st.session_state["cg_original_bytes"] = original_bytes
                log_app_event("document_analyzed")

            if "cg_results" in st.session_state:
                st.subheader("Suggestions de correction")
                results = st.session_state["cg_results"]
                st.write(f"{len(results)} suggestion(s)")
                if results:
                    st.dataframe(
                        results, use_container_width=True, hide_index=True
                    )
                else:
                    st.info(
                        "Aucune suggestion automatique trouvée. "
                        "Cela ne garantit pas l'absence d'erreurs."
                    )
        except Exception as exc:
            st.error(f"Impossible d'ouvrir le document : {exc}")
    else:
        st.info("Importez un document Word pour lancer les contrôles.")

    st.divider()
    st.subheader("Reformulation académique")
    rewrite_text = st.text_area(
        "Collez ici le passage à reformuler",
        height=180,
        key="correction_rewrite_text",
    )
    if st.button("Reformuler avec l'IA") and rewrite_text.strip():
        try:
            with st.spinner("Reformulation en cours..."):
                result = call_ai(
                    "Reformule en français académique sans inventer de faits, "
                    "de chiffres ou de sources. Conserve le sens du texte et "
                    "améliore la clarté :\n\n" + rewrite_text
                )
            st.text_area("Proposition à vérifier", result, height=240)
            st.session_state["cg_last_rewrite"] = result
            log_app_event("ai_rewrite")
        except Exception as exc:
            st.error(f"Erreur IA : {exc}")


# -------------------------------------------------------------------
# PAGE: RESEARCH AND WRITING
# -------------------------------------------------------------------
elif page == "Recherche & rédaction":
    st.header("Recherche & rédaction complète")
    st.info(
        "Préparez votre travail par étapes, recherchez des références réelles, "
        "puis exportez le contenu en Word et PDF. Sans clé IA active, le mode gratuit "
        "fournit une trame académique détaillée et des consignes de rédaction, pas un mémoire final rédigé automatiquement."
    )
    left, right = st.columns(2)
    with left:
        subject = st.text_input("Sujet / thème du travail", key="cg_rw_subject")
        work_type = st.selectbox(
            "Type de document",
            ["Mémoire universitaire", "Rapport de stage", "Exposé", "Article scientifique", "Projet de recherche"],
            key="cg_rw_type",
        )
        pages = st.slider("Nombre de pages souhaité", min_value=5, max_value=60, value=20, step=5, key="cg_rw_pages")
        language_out = st.selectbox("Langue de rédaction", ["Français", "Malagasy", "Anglais"], key="cg_rw_language")
    with right:
        level = st.selectbox("Niveau", ["Licence", "Master 1", "Master 2", "Doctorat", "Autre"], key="cg_rw_level")
        field = st.text_input("Filière / domaine", placeholder="Ex. Gestion commerciale, Commerce international")
        study_area = st.text_input("Terrain / zone d'étude", placeholder="Ex. Exportateurs de girofle à Toamasina")
        context = st.text_area(
            "Consignes, période et informations de terrain",
            height=145,
            placeholder="Décrivez l'établissement, l'entreprise, la période, les données disponibles et les consignes...",
            key="cg_rw_context",
        )

    st.caption(f"Estimation indicative : environ {pages * 250:,} mots pour {pages} pages, selon la mise en page et les exigences de votre établissement.".replace(",", " "))
    st.subheader("1. Recherche de livres et articles scientifiques")
    source_query = st.text_input(
        "Mots-clés de recherche documentaire",
        value=subject,
        key="cg_rw_source_query",
        placeholder="Saisissez votre sujet ou des mots-clés précis",
    )
    if st.button("Rechercher des sources réelles", key="cg_rw_search_sources"):
        if not source_query.strip():
            st.warning("Saisissez d'abord des mots-clés.")
        else:
            with st.spinner("Recherche dans Crossref et OpenAlex..."):
                st.session_state["cg_rw_sources"] = search_academic_sources(source_query.strip(), 6)
            if not st.session_state["cg_rw_sources"]:
                st.warning("Aucune référence trouvée ou les services sont temporairement indisponibles. Réessayez avec d'autres mots-clés.")
            else:
                st.success(f"{len(st.session_state['cg_rw_sources'])} référence(s) trouvée(s). Vérifiez leur pertinence avant citation.")

    sources = st.session_state.get("cg_rw_sources", [])
    if sources:
        for idx, source in enumerate(sources, 1):
            st.markdown(f"**{idx}. {source['title']}**")
            st.caption(f"{source['authors']} • {source['year'] or 'Année non indiquée'} • {source['source']}")
            if source.get("url"):
                st.markdown(f"[Consulter la source]({source['url']})")
            elif source.get("doi"):
                st.markdown(f"[Consulter le DOI](https://doi.org/{source['doi']})")

    st.divider()
    st.subheader("2. Rédaction par étapes")
    task = st.selectbox(
        "Que souhaitez-vous rédiger maintenant ?",
        [
            "Construire le plan détaillé complet",
            "Rédiger l'introduction générale",
            "Formuler la problématique, questions et hypothèses",
            "Rédiger le cadre théorique et la revue de littérature",
            "Rédiger le chapitre méthodologique",
            "Rédiger un chapitre d'analyse des résultats",
            "Rédiger la discussion et les recommandations",
            "Rédiger la conclusion générale",
            "Préparer la rédaction complète par étapes",
            "Améliorer un texte académique fourni",
        ],
        key="cg_rw_task",
    )
    chapter = st.text_input(
        "Titre ou sujet précis du chapitre (facultatif)",
        placeholder="Ex. Chapitre 2 : Les déterminants de la compétitivité à l'exportation",
        key="cg_rw_chapter",
    )
    supplied_text = st.text_area(
        "Texte à améliorer (uniquement pour la tâche d'amélioration)",
        height=120,
        key="cg_rw_supplied_text",
    )

    if st.button("Générer le contenu", type="primary", key="cg_rw_generate"):
        if not subject.strip():
            st.warning("Veuillez saisir le sujet du travail.")
        else:
            pages_for_section = max(2, pages // 5) if task == "Préparer la rédaction complète par étapes" else max(2, pages // 3)
            target_words = min(4500, pages_for_section * 250)
            prompt = f"""
Tu es un assistant de rédaction académique. Rédige en {language_out}, au niveau {level},
pour un document de type {work_type}, dans le domaine {field or 'à préciser'}.
Sujet : {subject}
Terrain : {study_area or 'à préciser'}
Consignes et données fournies : {context or 'Aucune donnée de terrain fournie'}
Tâche : {task}
Chapitre : {chapter or 'à définir selon le plan'}
Texte à améliorer : {supplied_text or 'Aucun texte fourni'}
Longueur cible pour cette section : environ {target_words} mots.
Exigences : structure claire, style universitaire, transitions, concepts définis,
aucun chiffre, résultat, auteur, citation ou référence inventé. N'affirme pas avoir
consulté une source si elle n'est pas dans les références fournies. Utilise des
marqueurs [SOURCE À AJOUTER] lorsque des preuves sont nécessaires. Distingue les
hypothèses des résultats réellement observés. Si les données manquent, indique
explicitement les éléments que l'étudiant doit collecter. Fournis un texte structuré
avec des titres hiérarchisés. Ne prétends pas que le nombre de pages est garanti.
"""
            api_key = get_secret("OPENAI_API_KEY")
            if api_key:
                try:
                    with st.spinner("Rédaction IA de la section en cours..."):
                        generated = call_ai(prompt)
                    st.session_state["cg_research_output"] = generated
                    st.success("Section générée. Relisez et vérifiez les informations et les références.")
                    log_app_event("research_writing_ai")
                except Exception as exc:
                    st.error(f"Rédaction IA indisponible : {exc}. Le mode gratuit reste disponible avec une trame.")
            else:
                if task == "Améliorer un texte académique fourni" and supplied_text.strip():
                    generated = (
                        "AIDE À L'AMÉLIORATION ACADÉMIQUE\n\n"
                        + offline_rewrite(supplied_text)
                        + "\n\nÀ vérifier : exactitude des faits, cohérence des arguments, citations et références."
                    )
                else:
                    generated = generate_free_research(
                        "Proposer un plan détaillé" if task in ("Construire le plan détaillé complet", "Préparer la rédaction complète par étapes") else task,
                        subject,
                        context + "\nType : " + work_type + "\nNiveau : " + level
                        + "\nFilière : " + field + "\nTerrain : " + study_area
                        + "\nNombre de pages visé : " + str(pages)
                        + "\nSection demandée : " + (chapter or task),
                        language_out,
                    )
                    generated = (
                        f"TRAME DE RÉDACTION — {task.upper()}\n"
                        f"Sujet : {subject}\nType : {work_type} • Niveau : {level} • Objectif : {pages} pages\n\n"
                        + generated
                        + "\n\nDÉVELOPPEMENT À COMPLÉTER\n"
                        "Pour atteindre le volume demandé, développez chaque sous-partie avec : "
                        "une définition appuyée par une source vérifiée, une explication, "
                        "un exemple lié au terrain, une donnée collectée et une courte synthèse.\n\n"
                        "ATTENTION : cette sortie hors ligne est une trame structurée, "
                        "pas un mémoire complet rédigé ni une preuve documentaire. "
                        "Ajoutez vos données réelles et vérifiez chaque référence."
                    )
                st.session_state["cg_research_output"] = generated
                st.info("Mode gratuit hors ligne : trame détaillée générée. La rédaction longue automatique nécessite une IA disponible.")
                log_app_event("research_writing_free")

    output = st.session_state.get("cg_research_output", "")
    if output:
        st.subheader("3. Résultat")
        edited_output = st.text_area("Relire et modifier le contenu", value=output, height=500, key="cg_rw_output_editor")
        st.session_state["cg_research_output"] = edited_output
        if st.button("Préparer les fichiers Word et PDF", key="cg_rw_prepare_exports"):
            try:
                document = make_research_docx(subject or "Travail académique", edited_output, sources)
                word_buffer = io.BytesIO()
                document.save(word_buffer)
                pdf_buffer = document_to_pdf_bytes(document, "Times New Roman", 12, "1.5", 2.5, True)
                st.session_state["cg_rw_docx"] = word_buffer.getvalue()
                st.session_state["cg_rw_pdf"] = pdf_buffer
                st.success("Les fichiers Word et PDF sont prêts.")
            except Exception as exc:
                st.error(f"Erreur de préparation des fichiers : {exc}")
        if st.session_state.get("cg_rw_docx"):
            st.download_button(
                "Télécharger la rédaction Word (.docx)",
                data=st.session_state["cg_rw_docx"],
                file_name="correctgloss_recherche_redaction.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="cg_rw_download_docx",
            )
        if st.session_state.get("cg_rw_pdf"):
            st.download_button(
                "Télécharger la rédaction PDF (.pdf)",
                data=st.session_state["cg_rw_pdf"],
                file_name="correctgloss_recherche_redaction.pdf",
                mime="application/pdf",
                key="cg_rw_download_pdf",
            )

    st.caption(
        "Les références repérées par Crossref/OpenAlex sont des métadonnées bibliographiques. "
        "Leur accès au texte intégral dépend de l'éditeur ou de la disponibilité en libre accès. "
        "Ne citez que les sources que vous avez effectivement consultées."
    )


# -------------------------------------------------------------------
# PAGE: EXPORT
# -------------------------------------------------------------------
elif page == "Export":
    st.header("Export Word et PDF")
    st.write("Importez un document Word, choisissez la mise en page, puis téléchargez les versions DOCX et PDF.")
    uploaded_export = st.file_uploader(
        "Importer le document Word (.docx)", type=["docx"], key="export_upload",
    )
    col1, col2 = st.columns(2)
    with col1:
        font = st.selectbox("Police", ["Times New Roman", "Arial", "Calibri"])
        font_size = st.number_input("Taille (pt)", 10, 14, 12)
        line_spacing = st.selectbox("Interligne", ["1.0", "1.15", "1.5", "2.0"], index=2)
    with col2:
        margin_cm = st.number_input("Marges (cm)", 1.5, 4.0, 2.5, 0.25)
        justify = st.checkbox("Justifier le texte", True)
        add_toc = st.checkbox("Ajouter une table des matières automatique (Word uniquement)", True)

    if uploaded_export and st.button("Préparer Word et PDF", type="primary"):
        try:
            formatted = Document(io.BytesIO(uploaded_export.getvalue()))
            for section in formatted.sections:
                section.top_margin = section.bottom_margin = Inches(margin_cm / 2.54)
                section.left_margin = section.right_margin = Inches(margin_cm / 2.54)
            for paragraph in formatted.paragraphs:
                if justify and not (
                    paragraph.style and paragraph.style.name.lower().startswith("heading")
                ):
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                paragraph.paragraph_format.line_spacing = float(line_spacing)
                paragraph.paragraph_format.space_after = Pt(6)
                for run in paragraph.runs:
                    run.font.name = font
                    run.font.size = Pt(int(font_size))
            if add_toc:
                formatted = add_word_toc(formatted)
            word_output = io.BytesIO()
            formatted.save(word_output)
            pdf_output = document_to_pdf_bytes(
                formatted, font, int(font_size), line_spacing, margin_cm, justify
            )
            st.session_state["cg_export_docx"] = word_output.getvalue()
            st.session_state["cg_export_pdf"] = pdf_output
            st.success("Les deux fichiers sont prêts au téléchargement.")
            log_app_event("docx_pdf_export_prepared")
        except Exception as exc:
            st.error(f"Erreur pendant la préparation des fichiers : {exc}")

    if st.session_state.get("cg_export_docx"):
        st.download_button(
            "Télécharger le Word (.docx)",
            data=st.session_state["cg_export_docx"],
            file_name="correctgloss_document_formate.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    if st.session_state.get("cg_export_pdf"):
        st.download_button(
            "Télécharger le PDF (.pdf)",
            data=st.session_state["cg_export_pdf"],
            file_name="correctgloss_document_formate.pdf",
            mime="application/pdf",
        )
    st.caption(
        "La table des matières automatique est intégrée dans le fichier Word. "
        "Dans Word, faites clic droit dessus puis « Mettre à jour les champs ». "
        "Le PDF reprend le texte, les titres et les tableaux avec une mise en page simple ; "
        "les images et certains éléments complexes du Word d'origine peuvent ne pas être reproduits."
    )


# -------------------------------------------------------------------
# PAGE: ADMIN DASHBOARD
# -------------------------------------------------------------------
elif page == "Dashboard administrateur":
    admin_dashboard()


st.divider()
st.caption(
    "CORRECTGLOSS — Les suggestions sont indicatives. "
    "Gardez toujours une copie originale de vos documents."
)
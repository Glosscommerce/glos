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
    """Generate structured academic scaffolds locally; never invent field data or references."""
    subject = subject.strip() or "[SUJET À PRÉCISER]"
    context = context.strip() or "[À compléter avec les informations vérifiées du terrain]"
    task_lower = task.lower()

    # Map labels used by the interface and the full-memory assistant to one template family.
    if "plan" in task_lower or "rédaction complète" in task_lower:
        family = "plan"
    elif "introduction" in task_lower:
        family = "introduction"
    elif "problématique" in task_lower or "hypoth" in task_lower or "question" in task_lower:
        family = "problem"
    elif "théorique" in task_lower or "littérature" in task_lower or "conceptuel" in task_lower or "définitions" in task_lower:
        family = "literature"
    elif "terrain" in task_lower or "organisation" in task_lower or "présentation du terrain" in task_lower:
        family = "terrain"
    elif "méthod" in task_lower:
        family = "methodology"
    elif "résultat" in task_lower or "analyse des données" in task_lower:
        family = "results"
    elif "discussion" in task_lower or "recommandation" in task_lower:
        family = "discussion"
    elif "conclusion" in task_lower:
        family = "conclusion"
    elif "amélior" in task_lower or "révision" in task_lower:
        family = "revision"
    else:
        family = "problem"

    if language == "Français":
        templates = {
            "plan": f"""PLAN DÉTAILLÉ PROPOSÉ — À ADAPTER
Sujet : {subject}

INTRODUCTION GÉNÉRALE
1. Contexte général et justification du sujet
2. Délimitation du sujet (espace, période, secteur)
3. Constat de départ et intérêt de la recherche
4. Problématique et question principale
5. Objectif général et objectifs spécifiques
6. Hypothèses de recherche
7. Méthodologie et sources de données
8. Annonce du plan

PREMIÈRE PARTIE — CADRE CONCEPTUEL ET THÉORIQUE
Chapitre 1 : Définitions, concepts clés et délimitation
  1.1 Définition des notions principales
  1.2 Dimensions et indicateurs des concepts
  1.3 Relations entre les concepts
Chapitre 2 : Théories et revue de littérature
  2.1 Théories et modèles explicatifs
  2.2 Travaux scientifiques antérieurs
  2.3 Synthèse critique et lacune de recherche

DEUXIÈME PARTIE — TERRAIN ET DÉMARCHE MÉTHODOLOGIQUE
Chapitre 3 : Présentation du terrain d'étude
  3.1 Historique et statut de l'organisation
  3.2 Activités, organisation et acteurs concernés
  3.3 Environnement et contexte de l'étude
Chapitre 4 : Méthodologie de recherche
  4.1 Type et approche de recherche
  4.2 Population, échantillon et critères de sélection
  4.3 Outils et procédure de collecte
  4.4 Méthodes de traitement et d'analyse
  4.5 Limites et considérations éthiques

TROISIÈME PARTIE — RÉSULTATS, DISCUSSION ET RECOMMANDATIONS
Chapitre 5 : Présentation et analyse des résultats réels
Chapitre 6 : Discussion, recommandations et plan d'action

CONCLUSION GÉNÉRALE
Réponse à la problématique, bilan des hypothèses, limites et perspectives.

Contexte fourni : {context}

À valider : ce plan est une proposition. Respectez le guide de mémoire de votre établissement et adaptez les parties à votre terrain réel.""",
            "introduction": f"""INTRODUCTION GÉNÉRALE — TRAME DE RÉDACTION
Sujet : {subject}

1. Contexte général
Présentez le domaine et son importance à l'aide de faits datés et de sources vérifiables.
[INSÉRER ICI LES FAITS ET SOURCES VÉRIFIÉS]

2. Justification et choix du sujet
Expliquez l'intérêt scientifique, économique, social ou professionnel du sujet.

3. Délimitation de l'étude
- Délimitation spatiale : [zone / entreprise / organisation]
- Délimitation temporelle : [période étudiée]
- Délimitation thématique : [aspects retenus et aspects exclus]

4. Constat et formulation du problème
Décrivez une situation réellement observée. Distinguez les faits, les opinions et les informations à vérifier.

5. Problématique
[EXPLIQUER LA DIFFICULTÉ CENTRALE ET LE MANQUE DE CONNAISSANCE À ÉTUDIER]

6. Question principale
Dans quelle mesure {subject.lower()} se manifeste-t-il dans le contexte étudié ?

7. Objectifs
Objectif général : analyser {subject.lower()} dans le terrain retenu.
Objectifs spécifiques :
- Décrire la situation et les pratiques existantes ;
- Identifier les facteurs explicatifs ;
- Analyser les informations collectées ;
- Proposer des améliorations fondées sur les résultats.

8. Hypothèses
[FORMULER DES HYPOTHÈSES QUI POURRONT ÊTRE VÉRIFIÉES PAR LES DONNÉES]

9. Méthodologie
Indiquez l'approche, la population étudiée, les outils de collecte et les méthodes d'analyse réellement utilisés.

10. Annonce du plan
Présentez brièvement les parties et chapitres du mémoire.

Informations de terrain à intégrer : {context}

Ne conservez aucun crochet dans la version finale sans le remplacer par une information vérifiée.""",
            "problem": f"""PROBLÉMATIQUE, QUESTIONS DE RECHERCHE ET HYPOTHÈSES
Sujet : {subject}

A. Constat de départ
[Décrire un fait observé, sa période, son lieu et la source qui permet de le vérifier.]

B. Problème de recherche
[Expliquer l'écart entre la situation actuelle et la situation souhaitée, ou le manque de connaissance.]

C. Question principale
Dans quelle mesure {subject.lower()} influence-t-il / se développe-t-il dans le terrain retenu ?

D. Questions spécifiques
1. Quelle est la situation actuelle du phénomène étudié ?
2. Quels facteurs économiques, organisationnels, techniques ou humains interviennent ?
3. Quelles conséquences et contraintes sont observées ?
4. Quelles solutions réalistes peuvent être proposées ?

E. Objectifs
Objectif général : analyser le phénomène étudié dans son contexte.
Objectifs spécifiques : décrire, identifier les facteurs, analyser les données et proposer des recommandations.

F. Hypothèses provisoires à tester
H1 : Les facteurs organisationnels et opérationnels ont une influence sur le phénomène étudié.
H2 : L'amélioration des pratiques et des ressources pourrait contribuer à de meilleurs résultats.
Ces hypothèses sont génériques : reformulez-les selon les variables et les données réellement accessibles.

Contexte et terrain fournis : {context}

Vérifiez que chaque hypothèse peut être confirmée ou rejetée à partir d'indicateurs observables.""",
            "literature": f"""CHAPITRE — CADRE CONCEPTUEL, THÉORIQUE ET REVUE DE LITTÉRATURE
Sujet : {subject}

1. Définition des concepts clés
Pour chaque concept :
- Donner une définition issue d'un auteur ou d'une institution identifiable ;
- Présenter la référence complète et la page si disponible ;
- Comparer les définitions ;
- Retenir une définition opérationnelle adaptée à l'étude.
[SOURCE SCIENTIFIQUE À RECHERCHER ET À VÉRIFIER]

2. Dimensions et indicateurs
Présenter les dimensions observables de chaque concept et les indicateurs qui permettront de les étudier.

3. Théories et modèles explicatifs
Exposer les théories pertinentes, leurs auteurs, leurs postulats et leurs limites.
[NE PAS ATTRIBUER UNE THÉORIE À UN AUTEUR SANS SOURCE VÉRIFIÉE]

4. Travaux empiriques antérieurs
Pour chaque étude : auteur, année, objectif, terrain, méthode, résultats publiés et limites.
[AJOUTER UNIQUEMENT DES ÉTUDES EFFECTIVEMENT CONSULTÉES]

5. Synthèse critique
Comparer les résultats des auteurs, relever les convergences et divergences et identifier la lacune à laquelle répond le présent mémoire.

Contexte fourni : {context}

Cette trame ne constitue pas une bibliographie. Recherchez et lisez les documents avant de citer leurs auteurs.""",
            "terrain": f"""CHAPITRE — PRÉSENTATION DU TERRAIN D'ÉTUDE
Sujet : {subject}

1. Identification de l'organisation ou de la zone
- Dénomination officielle : [À FOURNIR]
- Localisation : [À FOURNIR]
- Statut juridique / institutionnel : [À VÉRIFIER]
- Date de création et historique : [SOURCE OU DOCUMENT À FOURNIR]

2. Missions et activités
Décrire uniquement les activités confirmées par des documents ou des entretiens autorisés.

3. Organisation et fonctionnement
Présenter l'organigramme, les services, les responsabilités et les circuits de travail à partir d'informations validées.

4. Environnement de l'étude
Décrire les acteurs, partenaires, clients, fournisseurs, concurrents ou institutions concernés selon le sujet.

5. Lien entre le terrain et la problématique
Expliquer pourquoi ce terrain permet d'étudier la question de recherche.

Informations réellement fournies : {context}

Ne pas inventer l'historique, les effectifs, le chiffre d'affaires, l'organigramme ou les activités de l'organisation.""",
            "methodology": f"""CHAPITRE — MÉTHODOLOGIE DE RECHERCHE
Sujet : {subject}

1. Objectif et logique de la démarche
Expliquer comment la méthode choisie permet de répondre aux questions de recherche.

2. Approche de recherche
[QUALITATIVE / QUANTITATIVE / MIXTE — JUSTIFIER LE CHOIX]

3. Population et unité d'analyse
- Population cible : [À DÉFINIR]
- Unité d'analyse : [PERSONNE / ENTREPRISE / DOCUMENT / TRANSACTION...]
- Critères d'inclusion et d'exclusion : [À PRÉCISER]

4. Échantillonnage
- Méthode : [ALÉATOIRE / RAISONNÉ / EXHAUSTIF / AUTRE]
- Taille de l'échantillon : [À RENSEIGNER APRÈS DÉFINITION]
- Justification : [À EXPLIQUER]

5. Techniques et outils de collecte
Questionnaire, entretien, observation, analyse documentaire ou base de données : préciser ce qui a réellement été utilisé et joindre les outils en annexe.

6. Traitement et analyse
Décrire les opérations de codage, classement, calcul ou analyse thématique réellement effectuées.

7. Fiabilité, validité et éthique
Expliquer le contrôle des informations, le consentement, la confidentialité et les limites de la démarche.

Contexte fourni : {context}

Ne déclarez pas qu'une enquête, un entretien ou une observation a été réalisé si ce n'est pas le cas.""",
            "results": f"""CHAPITRE — PRÉSENTATION ET ANALYSE DES RÉSULTATS
Sujet : {subject}

IMPORTANT : aucun résultat ne peut être produit sans données réelles. Utilisez la structure ci-dessous après la collecte.

1. Rappel de l'objectif et de la méthode
[PRÉCISER LA PÉRIODE, LES OUTILS ET LE NOMBRE RÉEL DE RÉPONSES]

2. Présentation des données
Tableau 1 : [INDICATEUR] | [VALEUR OBSERVÉE] | [SOURCE / PÉRIODE]
Tableau 2 : [INDICATEUR] | [VALEUR OBSERVÉE] | [SOURCE / PÉRIODE]

3. Analyse descriptive
Décrire les tendances visibles sans extrapoler au-delà de l'échantillon.

4. Analyse par question ou variable
Pour chaque question : présenter les données, expliquer ce qu'elles indiquent et signaler les limites.

5. Synthèse des constats
Distinguer les résultats confirmés, les résultats incertains et les informations manquantes.

Données fournies : {context}

[DONNÉES DE TERRAIN À FOURNIR]
N'inventez aucun pourcentage, effectif, verbatim d'entretien ou résultat statistique.""",
            "discussion": f"""CHAPITRE — DISCUSSION ET RECOMMANDATIONS
Sujet : {subject}

1. Rappel des principaux résultats
[REPRENDRE UNIQUEMENT LES RÉSULTATS ÉTABLIS DANS LE CHAPITRE PRÉCÉDENT]

2. Interprétation
Expliquer les résultats en lien avec la problématique, les objectifs et les variables.

3. Comparaison avec la littérature
Comparer les résultats aux travaux réellement consultés. Expliquer les accords et les différences avec prudence.
[SOURCE SCIENTIFIQUE À VÉRIFIER]

4. Vérification des hypothèses
Pour chaque hypothèse : indicateurs examinés, éléments favorables, éléments contraires et conclusion justifiée (confirmée, partiellement confirmée ou rejetée).

5. Recommandations opérationnelles
Pour chaque recommandation, préciser :
- Problème ciblé ;
- Action proposée ;
- Responsable potentiel ;
- Ressources nécessaires ;
- Délai ;
- Indicateur de suivi.

6. Limites et perspectives
Présenter les limites des données, de l'échantillon, de la période et de la méthode.

Contexte fourni : {context}

Les recommandations doivent découler des résultats observés et non de suppositions.""",
            "conclusion": f"""CONCLUSION GÉNÉRALE — TRAME À COMPLÉTER
Sujet : {subject}

1. Rappeler le contexte, la problématique et l'objectif général.
2. Résumer brièvement la démarche méthodologique réellement appliquée.
3. Présenter les principaux résultats vérifiés, sans ajouter de nouvelle donnée.
4. Répondre directement à la question principale.
5. Faire le bilan de chaque hypothèse à partir des résultats.
6. Présenter les recommandations prioritaires.
7. Reconnaître les limites de l'étude.
8. Ouvrir sur une perspective de recherche ou d'action.

Résultats et informations fournis : {context}

La conclusion finale doit être rédigée après l'analyse des résultats. Ne prétendez pas répondre à la problématique si les données ne sont pas encore disponibles.""",
            "revision": f"""AIDE À LA RÉVISION ACADÉMIQUE HORS LIGNE
Sujet : {subject}

TEXTE FOURNI :
{context}

CORRECTIONS AUTOMATIQUES LIMITÉES
Le système peut corriger quelques formulations répétitives et espaces, mais ne remplace pas une correction grammaticale complète.

GRILLE DE RÉVISION
1. Une idée principale clairement formulée par paragraphe.
2. Des arguments appuyés par des preuves ou des sources vérifiables.
3. Des transitions logiques entre les paragraphes.
4. Des termes précis et cohérents avec le domaine.
5. Une distinction claire entre faits, interprétations et opinions.
6. Des citations et références complètes et vérifiables.

PROPOSITIONS DE STYLE
- Remplacer les formulations vagues par des termes précis.
- Éviter les répétitions et les phrases trop longues.
- Ne pas transformer une hypothèse en résultat.
- Conserver le sens original de l'auteur.

Texte révisé mécaniquement :
{offline_rewrite(context)}

Relisez chaque modification : une reformulation automatique simple peut changer le sens d'une phrase."""
        }
        result = templates[family]
        if family not in ("revision",):
            result += "\n\n" + "FICHE DE CONTRÔLE AVANT VALIDATION\n- Vérifier les faits et les dates.\n- Ajouter les sources réellement consultées.\n- Remplacer les champs entre crochets.\n- Vérifier la cohérence avec la problématique et les consignes de l'établissement."
        return result

    return f"""OFFLINE ACADEMIC WORKSHEET
Topic: {subject}
Selected language: {language}
Requested section: {task}

1. Background and verified context: [ADD VERIFIED FACTS AND SOURCES]
2. Research problem: [DEFINE THE SPECIFIC PROBLEM]
3. Main and specific research questions: [COMPLETE]
4. General and specific objectives: [COMPLETE]
5. Working hypotheses and measurable indicators: [COMPLETE IF REQUIRED]
6. Study area, population and methodology: [ADD REAL INFORMATION]
7. Findings: [INSERT ONLY COLLECTED AND VERIFIED DATA]
8. Discussion, recommendations and conclusion: [COMPLETE AFTER ANALYSIS]

Context supplied:
{context}

This is an offline worksheet, not an automatically written academic chapter. Translate and adapt it into {language}; do not invent facts, results, authors or references."""

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
    """Apply conservative, transparent local cleanups without an AI service."""
    cleaned = re.sub(r"[ \t]+", " ", text.strip())
    cleaned = re.sub(r"\n{3,}", "\n\n", cleaned)
    for pattern, replacement in [
        (r"\bau jour d'aujourd'hui\b", "aujourd'hui"),
        (r"\bmalgré que\b", "bien que"),
        (r"\bon peut voir que\b", "les résultats montrent que"),
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
        "Mode gratuit hors ligne : construisez votre plan, problématique, hypothèses, méthodologie et chapitres à partir de trames détaillées. "
        "Aucun crédit OpenAI n'est nécessaire. Les données de terrain, résultats et références doivent être fournis et vérifiés par l'utilisateur."
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


    # Assistant de mémoire complet : génération et sauvegarde section par section.
    with st.expander("Assistant mémoire complet — construire le document étape par étape", expanded=True):
        st.write(
            "Rédigez chaque partie séparément. Les sections validées sont conservées "
            "pendant la session et peuvent être assemblées dans un seul document."
        )
        full_stages = [
            "Introduction générale",
            "Chapitre 1 — Cadre conceptuel et définitions",
            "Chapitre 2 — Cadre théorique et revue de littérature",
            "Chapitre 3 — Présentation du terrain d'étude",
            "Chapitre 4 — Méthodologie de recherche",
            "Chapitre 5 — Présentation et analyse des résultats",
            "Chapitre 6 — Discussion et recommandations",
            "Conclusion générale",
        ]
        selected_stage = st.selectbox(
            "Partie du mémoire à rédiger", full_stages, key="cg_full_stage"
        )
        stage_details = st.text_area(
            "Précisions propres à cette partie (facultatif)",
            placeholder="Ex. données disponibles, consignes de l'encadreur, période étudiée...",
            key="cg_full_stage_details",
            height=90,
        )
        if "cg_full_sections" not in st.session_state:
            st.session_state["cg_full_sections"] = {}
        if st.button("Générer cette partie", type="primary", key="cg_full_generate"):
            if not subject.strip():
                st.warning("Saisissez d'abord le sujet du mémoire.")
            else:
                full_prompt = f"""
Tu es un assistant de rédaction universitaire. Rédige uniquement la partie suivante :
{selected_stage}
Sujet : {subject}
Type : {work_type}; niveau : {level}; domaine : {field or 'à préciser'}
Terrain : {study_area or 'à préciser'}
Langue : {language_out}
Contexte général et données réellement fournies : {context or 'Aucune donnée de terrain fournie'}
Précisions sur cette partie : {stage_details or 'Aucune'}
Longueur indicative : {max(500, int(pages * 250 / len(full_stages)))} mots.
Consignes :
- Structure avec titres et sous-titres cohérents et rédaction académique.
- Ne fabrique aucun chiffre, résultat d'enquête, entretien, auteur, citation ou référence.
- Pour les faits à documenter, indique [SOURCE À VÉRIFIER].
- Pour les résultats empiriques absents, indique clairement [DONNÉES DE TERRAIN À FOURNIR].
- N'affirme jamais qu'une hypothèse est validée sans données.
- Assure la cohérence avec le sujet, le niveau et le terrain.
"""
                stage_context = "\n".join([
                    "Type : " + work_type,
                    "Niveau : " + level,
                    "Domaine : " + (field or "[À préciser]"),
                    "Terrain : " + (study_area or "[À préciser]"),
                    "Contexte : " + (context or "[À compléter]"),
                    "Précisions : " + (stage_details or "[À compléter]"),
                    "Volume cible : " + str(pages) + " pages pour l'ensemble du document",
                ])
                st.session_state["cg_full_sections"][selected_stage] = generate_free_research(
                    selected_stage, subject, stage_context, language_out
                )
                st.success("Trame structurée créée hors ligne et sauvegardée. Complétez-la avec vos informations vérifiées.")
                st.info("Aucune rédaction IA payante n'est appelée dans ce module. Les chapitres de résultats restent à compléter avec vos données réelles.")

        saved_sections = st.session_state["cg_full_sections"]
        st.caption(f"Parties sauvegardées : {len(saved_sections)} / {len(full_stages)}")
        if selected_stage in saved_sections:
            edited_section = st.text_area(
                "Relire et modifier la partie sélectionnée",
                value=saved_sections[selected_stage],
                height=280,
                key="cg_full_editor_" + str(full_stages.index(selected_stage)),
            )
            if st.button("Enregistrer mes modifications", key="cg_full_save_section"):
                saved_sections[selected_stage] = edited_section
                st.session_state["cg_full_sections"] = saved_sections
                st.success("Modifications enregistrées.")
        if saved_sections:
            with st.expander("Voir les parties déjà sauvegardées"):
                for saved_title, saved_body in saved_sections.items():
                    st.markdown(f"**{saved_title}**")
                    st.caption(f"{len(saved_body.split())} mots")
            if st.button("Assembler le mémoire en Word et PDF", key="cg_full_assemble"):
                assembled = "\n\n".join(
                    title.upper() + "\n\n" + saved_sections[title]
                    for title in full_stages if title in saved_sections
                )
                assembled_doc = make_research_docx(subject or "Mémoire", assembled, sources)
                docx_io = io.BytesIO()
                assembled_doc.save(docx_io)
                st.session_state["cg_full_docx"] = docx_io.getvalue()
                st.session_state["cg_full_pdf"] = document_to_pdf_bytes(
                    assembled_doc, "Times New Roman", 12, "1.5", 2.5, True
                )
                st.success("Mémoire assemblé. Vérifiez les titres, les références et les données avant utilisation.")
        if st.session_state.get("cg_full_docx"):
            st.download_button(
                "Télécharger le mémoire complet (.docx)",
                data=st.session_state["cg_full_docx"],
                file_name="correctgloss_memoire_complet.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="cg_full_download_docx",
            )
        if st.session_state.get("cg_full_pdf"):
            st.download_button(
                "Télécharger le mémoire complet (.pdf)",
                data=st.session_state["cg_full_pdf"],
                file_name="correctgloss_memoire_complet.pdf",
                mime="application/pdf",
                key="cg_full_download_pdf",
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
            if task == "Améliorer un texte académique fourni" and supplied_text.strip():
                generated = generate_free_research(
                    "Améliorer un paragraphe académique",
                    subject,
                    supplied_text,
                    language_out,
                )
            else:
                generated = generate_free_research(
                    task,
                    subject,
                    "\n".join([
                        context,
                        "Type : " + work_type,
                        "Niveau : " + level,
                        "Filière : " + field,
                        "Terrain : " + study_area,
                        "Nombre de pages visé : " + str(pages),
                        "Section demandée : " + (chapter or task),
                    ]),
                    language_out,
                )
                generated = (
                    f"TRAME HORS LIGNE — {task.upper()}\n"
                    f"Sujet : {subject}\nType : {work_type} • Niveau : {level} • Volume visé : {pages} pages\n\n"
                    + generated
                    + "\n\nDÉVELOPPEMENT À COMPLÉTER\n"
                    "Développez chaque sous-partie avec une explication, une source réellement consultée, "
                    "un exemple pertinent et, lorsque nécessaire, une donnée collectée sur le terrain."
                )
            st.session_state["cg_research_output"] = generated
            st.info("Mode gratuit hors ligne activé : aucun appel à l'API OpenAI n'est effectué pour la Recherche & rédaction.")
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
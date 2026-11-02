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
    st.header("Recherche & rédaction")
    st.info(
        "Mode gratuit hors ligne : les propositions utilisent des modèles académiques intégrés. "
        "Aucune clé API ni aucun crédit OpenAI ne sont nécessaires. Vérifiez toujours les faits, "
        "chiffres et références."
    )
    task = st.selectbox(
        "Type de travail",
        [
            "Proposer un plan détaillé",
            "Rédiger une introduction",
            "Formuler une problématique et des hypothèses",
            "Améliorer un paragraphe académique",
            "Préparer une conclusion",
            "Créer des questions de recherche",
        ],
    )
    subject = st.text_input("Sujet / thème du travail")
    context = st.text_area(
        "Contexte, consignes ou texte de départ",
        height=180,
        placeholder="Précisez le niveau, la filière, le terrain d'étude et les consignes...",
    )
    language_out = st.selectbox(
        "Langue de rédaction",
        ["Français", "Anglais", "Malagasy"],
    )
    if st.button("Générer une proposition", type="primary"):
        if not subject.strip():
            st.warning("Veuillez saisir le sujet de votre travail.")
        else:
            try:
                with st.spinner("Préparation du contenu hors ligne..."):
                    output = generate_free_research(task, subject, context, language_out)
                st.session_state["cg_research_output"] = output
                st.text_area("Résultat à relire et compléter", output, height=420)
                st.success("Proposition créée en mode gratuit, sans clé API ni crédit OpenAI.")
                log_app_event("research_writing_free")
            except Exception as exc:
                st.error(f"Erreur lors de la génération hors ligne : {exc}")

    if st.session_state.get("cg_research_output"):
        st.download_button(
            "Télécharger le texte (.txt)",
            data=st.session_state["cg_research_output"],
            file_name="correctgloss_redaction.txt",
            mime="text/plain",
        )


# -------------------------------------------------------------------
# PAGE: EXPORT
# -------------------------------------------------------------------
elif page == "Export":
    st.header("Export et mise en page Word")
    st.write("Importez un fichier Word, puis choisissez les paramètres de mise en page.")
    uploaded_export = st.file_uploader(
        "Importer le document Word (.docx)",
        type=["docx"],
        key="export_upload",
    )
    col1, col2 = st.columns(2)
    with col1:
        font = st.selectbox("Police", ["Times New Roman", "Arial", "Calibri"])
        font_size = st.number_input("Taille (pt)", 10, 14, 12)
        line_spacing = st.selectbox(
            "Interligne", ["1.0", "1.15", "1.5", "2.0"], index=2
        )
    with col2:
        margin_cm = st.number_input("Marges (cm)", 1.5, 4.0, 2.5, 0.25)
        justify = st.checkbox("Justifier le texte", True)
        add_toc = st.checkbox("Ajouter une table des matières automatique", True)

    if uploaded_export and st.button("Préparer le DOCX", type="primary"):
        try:
            formatted = Document(io.BytesIO(uploaded_export.getvalue()))
            for section in formatted.sections:
                section.top_margin = section.bottom_margin = Inches(margin_cm / 2.54)
                section.left_margin = section.right_margin = Inches(margin_cm / 2.54)
            for paragraph in formatted.paragraphs:
                if justify and not (
                    paragraph.style
                    and paragraph.style.name.lower().startswith("heading")
                ):
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                paragraph.paragraph_format.line_spacing = float(line_spacing)
                paragraph.paragraph_format.space_after = Pt(6)
                for run in paragraph.runs:
                    run.font.name = font
                    run.font.size = Pt(int(font_size))
            if add_toc:
                formatted = add_word_toc(formatted)
            output = io.BytesIO()
            formatted.save(output)
            st.session_state["cg_export_docx"] = output.getvalue()
            st.success(
                "Document Word préparé. Dans Word, faites clic droit sur la "
                "table des matières puis « Mettre à jour les champs »."
            )
            log_app_event("docx_export_prepared")
        except Exception as exc:
            st.error(f"Erreur pendant la préparation du document : {exc}")

    if st.session_state.get("cg_export_docx"):
        st.download_button(
            "Télécharger le DOCX",
            data=st.session_state["cg_export_docx"],
            file_name="correctgloss_document_formate.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    st.caption(
        "Cette version exporte en Word (.docx). L'export PDF sera ajouté dans "
        "une prochaine étape après vérification des dépendances."
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
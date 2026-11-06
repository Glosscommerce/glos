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
    """Create a substantive editable starter draft without inventing research evidence."""
    fr = lang == "Français"
    subject = re.sub(r"\s+", " ", str(subject or "")).strip(" .;:")
    clove = any(x in subject.lower() for x in ("girofle", "clove"))
    if fr and kind == "Mémoire / Projet de recherche":
        if clove:
            context = "Le girofle représente un produit agricole commercialisé sur les marchés internationaux. Dans le contexte malgache, son étude présente un intérêt pour comprendre les relations entre production, collecte, préparation, contrôle de la qualité et exportation. La qualité peut influencer l’acceptation des lots, la satisfaction des acheteurs et la continuité des relations commerciales. Ces éléments devront être documentés à l’aide de statistiques récentes et de sources professionnelles vérifiables."
            scope = "Le présent projet porte sur « " + subject + " ». Le terrain envisagé est constitué des exportateurs de girofle établis à Toamasina. L’étude examinera les critères de qualité appliqués au produit destiné à l’exportation, les pratiques de contrôle et de conservation, ainsi que les contraintes rencontrées par les opérateurs. La période étudiée et les entreprises participantes seront précisées après vérification de l’accès au terrain."
            problem = "La compétitivité à l’exportation ne dépend pas uniquement du prix. Elle peut également être liée à la conformité du produit aux spécifications des acheteurs, à la régularité de la qualité, à la traçabilité, au conditionnement et au respect des délais. La question centrale est donc la suivante : dans quelle mesure la qualité du girofle malgache influence-t-elle la compétitivité des exportateurs de Toamasina sur les marchés internationaux ?"
            questions = ["Quels critères de qualité les acheteurs exigent-ils pour le girofle destiné à l’exportation ?", "Quelles pratiques les exportateurs de Toamasina utilisent-ils pour contrôler et préserver cette qualité ?", "Quelles contraintes limitent la conformité du produit aux exigences des marchés ?", "Comment l’amélioration de la gestion de la qualité pourrait-elle renforcer la compétitivité des exportateurs ?"]
            hypotheses = ["H1 : Le respect régulier des critères de qualité exigés par les acheteurs favorise le maintien des relations commerciales.", "H2 : Les insuffisances de contrôle, de stockage ou de traçabilité peuvent réduire la conformité des lots.", "H3 : Le renforcement des procédures de contrôle qualité pourrait améliorer la performance commerciale. Ces hypothèses restent à vérifier sur le terrain."]
            theory = ["La qualité sera étudiée comme l’ensemble des caractéristiques permettant au produit de répondre à des exigences définies. Les dimensions à examiner pourront inclure l’état physique, la propreté, l’humidité, l’homogénéité, le stockage, le conditionnement et la conformité aux spécifications contractuelles. Les seuils précis devront provenir de normes ou de documents techniques vérifiés.", "La compétitivité à l’exportation désigne la capacité d’un opérateur à maintenir durablement ses ventes sur les marchés extérieurs. Elle pourra être examinée à travers la conformité des livraisons, la fidélisation des clients, la régularité des commandes, les prix obtenus et les volumes exportés, selon les données réellement disponibles.", "La gestion de la qualité, la chaîne de valeur et les théories du commerce international pourront servir de cadres d’analyse. Les auteurs et modèles retenus seront présentés après consultation d’ouvrages et d’articles scientifiques réellement accessibles et vérifiés.", "Le terrain de Toamasina sera décrit à partir de documents officiels et d’informations recueillies auprès des acteurs concernés. Les volumes, destinations et caractéristiques de la filière ne seront pas estimés sans source."]
        else:
            context = "Le sujet « " + subject + " » s’inscrit dans un contexte qui mérite une analyse scientifique afin de comprendre le phénomène étudié, ses facteurs explicatifs et ses conséquences. L’intérêt de la recherche est de produire une analyse structurée, fondée sur des concepts clairement définis et des informations vérifiables."
            scope = "L’étude devra être délimitée selon trois dimensions : le lieu, la période et la population concernée. Ces limites permettront de préciser les observations possibles et d’éviter des conclusions trop générales. Le terrain et les participants seront décrits à partir des informations effectivement obtenues."
            problem = "La problématique consiste à comprendre les mécanismes qui relient les différentes dimensions du sujet « " + subject + " ». La question centrale devra être formulée de manière précise, neutre et vérifiable ; elle guidera le choix des concepts, des variables, des outils de collecte et des méthodes d’analyse."
            questions = ["Quelle est la situation actuelle du phénomène étudié sur le terrain retenu ?", "Quels facteurs peuvent expliquer les évolutions ou difficultés observées ?", "Quelles conséquences ces facteurs ont-ils pour les acteurs concernés ?", "Quelles actions pourraient être proposées à partir des résultats ?"]
            hypotheses = ["H1 : Le phénomène étudié est associé à plusieurs facteurs qu’il faudra identifier sur le terrain.", "H2 : Les pratiques des acteurs et les ressources disponibles peuvent influencer les résultats observés.", "H3 : Des mesures adaptées au contexte pourraient contribuer à améliorer la situation. Ces hypothèses sont provisoires."]
            theory = ["Le cadre conceptuel définira les notions centrales à partir de références scientifiques et institutionnelles, puis précisera leurs dimensions observables.", "Le cadre théorique présentera les approches permettant d’interpréter le phénomène. Chaque théorie devra être justifiée par son lien avec la problématique.", "La revue des travaux antérieurs permettra de comparer les connaissances disponibles et d’identifier la contribution attendue de cette recherche."]
        sections = [
            ("INTRODUCTION GÉNÉRALE", [context, scope, problem, "Questions de recherche : " + " ".join(questions), "Hypothèses provisoires : " + " ".join(hypotheses), "Objectif général : analyser le sujet dans le terrain retenu. Objectifs spécifiques : définir les concepts, décrire les pratiques des acteurs, identifier les facteurs explicatifs, analyser les données et formuler des recommandations.", "Annonce du plan : la première partie présente le cadre conceptuel et théorique ; la deuxième expose la méthodologie ; la troisième présente et discute les résultats."]),
            ("PARTIE I — CADRE CONCEPTUEL ET THÉORIQUE", theory + ["La revue de littérature sera organisée par concepts, théories et résultats de recherches antérieures. Toute idée empruntée devra être accompagnée d’une citation vérifiable.", "La présentation du secteur ou du terrain précisera les acteurs, les activités et les caractéristiques utiles, à partir de documents fiables."]),
            ("PARTIE II — CADRE MÉTHODOLOGIQUE", ["La recherche adoptera une démarche descriptive et analytique. Une approche qualitative, quantitative ou mixte sera choisie selon la question de recherche et l’accès réel aux données.", "La population sera constituée des acteurs directement concernés. L’échantillon sera défini selon des critères explicites : fonction, expérience, disponibilité et lien avec le sujet. Le nombre de participants sera fixé après évaluation de l’accès au terrain.", "Les données primaires pourront être recueillies par entretiens semi-directifs, questionnaires et observations. Les données secondaires pourront provenir de rapports, statistiques, textes officiels et publications dont l’origine et la date auront été vérifiées.", "Les réponses qualitatives seront classées par thèmes. Les données quantitatives disponibles seront présentées dans des tableaux et indicateurs adaptés. Les limites possibles concernent l’accès au terrain, la disponibilité des données, la taille de l’échantillon et la confidentialité."]),
            ("PARTIE III — RÉSULTATS, ANALYSE ET DISCUSSION", ["Cette partie sera rédigée après la collecte. Elle présentera les participants, les données recueillies et les résultats organisés selon les objectifs spécifiques.", "L’analyse reliera chaque observation à la problématique et au cadre théorique. Toute interprétation devra s’appuyer sur un entretien, un questionnaire, une observation ou un document vérifiable.", "La discussion déterminera si les résultats confirment, nuancent ou rejettent les hypothèses. Aucun résultat empirique ne doit être inventé.", "Les recommandations découleront des difficultés réellement constatées et préciseront les acteurs concernés ainsi que les actions envisageables."]),
            ("CONCLUSION GÉNÉRALE", ["La conclusion rappellera la problématique, les objectifs et la démarche, puis présentera les principaux résultats effectivement établis et répondra à la question centrale.", "Les limites de l’étude seront exposées avec transparence : terrain, échantillon, période et disponibilité des données.", "Les perspectives pourront proposer un élargissement du terrain, une comparaison avec d’autres acteurs ou l’approfondissement d’une dimension insuffisamment documentée."]),
            ("BIBLIOGRAPHIE ET ANNEXES", ["La bibliographie comprendra uniquement les documents réellement consultés : ouvrages, articles scientifiques, rapports institutionnels, textes officiels et ressources en ligne. Chaque référence devra être vérifiable et respecter la norme exigée par l’établissement.", "Les annexes pourront contenir le questionnaire, le guide d’entretien, les tableaux autorisés et les documents utiles. Les informations confidentielles devront être anonymisées ou utilisées avec autorisation."])
        ]
    elif fr:
        sections = [("INTRODUCTION", ["Le présent document porte sur « " + subject + " ». Il vise à présenter le sujet, son contexte et son intérêt. Les informations factuelles devront être appuyées par des sources vérifiées.", "La réflexion sera organisée autour des idées principales et des objectifs du document."]), ("DÉVELOPPEMENT", ["Premièrement, les notions essentielles et le contexte seront présentés à partir de références fiables.", "Deuxièmement, les arguments seront développés et illustrés par des exemples vérifiables.", "Enfin, les observations seront discutées en tenant compte des limites des informations disponibles."]), ("CONCLUSION", ["La conclusion synthétisera les arguments et formulera une réponse cohérente avec les éléments présentés."]), ("SOURCES", ["Ajouter uniquement les références effectivement consultées et vérifiées."])]
    else:
        sections = [("INTRODUCTION", ["This document examines “" + subject + "”, introduces its context and explains its relevance. Claims must be supported by verified sources."]), ("MAIN DISCUSSION", ["Define the key concepts using reliable sources.", "Explain the main factors and compare relevant perspectives.", "Present findings only after collecting and checking evidence."]), ("METHODOLOGY", ["Describe the research design, population, sampling criteria, data collection tools and analysis methods. Do not report fieldwork that has not taken place."]), ("CONCLUSION", ["Summarize verified findings, answer the research question and state the study limitations."]), ("REFERENCES", ["Include only sources actually consulted and verified."])]
    if instructions.strip():
        sections.insert(0, ("CONSIGNES PARTICULIÈRES" if fr else "SPECIAL INSTRUCTIONS", [instructions.strip()]))
    if not full_draft:
        sections.append(("QUESTIONS POUR VOUS GUIDER", ["Quelle question précise le document traite-t-il ?", "Quelles informations sont vérifiées et quelles données restent à collecter ?", "Les conclusions répondent-elles aux objectifs annoncés ?"] if fr else ["What specific question does the document address?", "Which information is verified and what data remains to be collected?", "Do the conclusions answer the stated objectives?"]))
    return [("TITLE", subject), ("Type", kind), ("Level", level), ("Language", lang)] + sections

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
            prepared_doc.add_heading(str(heading), level=1)
            body_paragraphs = (
                list(body) if isinstance(body, (list, tuple))
                else str(body or "").splitlines()
            )
            body_text_parts = []
            for paragraph_text in body_paragraphs:
                paragraph_text = str(paragraph_text).strip()
                if paragraph_text:
                    prepared_doc.add_paragraph(paragraph_text)
                    body_text_parts.append(paragraph_text)
            prepared_text_parts.extend([str(heading), "\n".join(body_text_parts)])
        prepared_buffer = io.BytesIO()
        prepared_doc.save(prepared_buffer)
        prepared_docx_bytes = prepared_buffer.getvalue()
        st.session_state["cg_prepared_writing_docx"] = prepared_docx_bytes
        try:
            st.session_state["cg_prepared_writing_pdf"] = build_pdf_from_docx(
                prepared_docx_bytes, font="Times New Roman", font_size=12,
                line_spacing="1.5", margin_cm=2.5,
            )
        except Exception:
            st.session_state.pop("cg_prepared_writing_pdf", None)
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
    writing_filename = st.session_state.get("cg_prepared_writing_name", "redaction")
    st.download_button(
        "Télécharger la rédaction préparée (Word)",
        st.session_state["cg_prepared_writing_docx"],
        file_name=f"correctgloss_{writing_filename}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    if "cg_prepared_writing_pdf" in st.session_state:
        st.download_button(
            "Télécharger la rédaction préparée (PDF)",
            st.session_state["cg_prepared_writing_pdf"],
            file_name=f"correctgloss_{writing_filename}.pdf",
            mime="application/pdf",
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

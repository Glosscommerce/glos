"""CORRECTGLOSS V3 — separate test application; never replaces V2 app.py."""
from __future__ import annotations

import os
from io import BytesIO

import streamlit as st
from docx import Document

try:
    for key in (
        "AI_PROVIDER", "HF_TOKEN", "HF_MODEL", "HF_CHAT_COMPLETIONS_URL",
        "AI_MAX_TOKENS", "OLLAMA_BASE_URL", "OLLAMA_MODEL",
    ):
        value = st.secrets.get(key)
        if value is not None and str(value).strip():
            os.environ[key] = str(value).strip()
except Exception:
    pass

from v3_free.ai_provider import (  # noqa: E402
    AIConfigurationError, AIProviderError, provider_status,
)
from v3_free.assistant import answer_question  # noqa: E402
from v3_free.long_writer import create_outline, write_section  # noqa: E402
from v3_free.document_export import build_docx, build_pdf  # noqa: E402
from correction_engine import (  # noqa: E402
    analyze_document_structure, correct_docx_with_languagetool,
)
from export_engine import build_pdf_from_docx  # noqa: E402

st.set_page_config(
    page_title="CORRECTGLOSS V3",
    page_icon="📘",
    layout="wide",
)

st.title("📘 CORRECTGLOSS V3")
st.caption("Assistant intelligent et rédaction académique — environnement de test séparé de V2")

with st.sidebar:
    st.header("⚙️ Configuration")
    language = st.selectbox("Langue de réponse", ["Français", "Malagasy", "English"])
    status = provider_status()
    st.markdown("**Moteur IA :** " + status["provider"])
    if status["configured"] == "no":
        st.warning(
            "Moteur IA non configuré. Ajoutez HF_TOKEN et HF_MODEL dans les secrets "
            "de cette application de test, ou configurez un serveur Ollama accessible."
        )
    else:
        st.info(status["note"])
    st.caption(
        "Les quotas gratuits dépendent du fournisseur. Les conversations et brouillons "
        "restent dans la session du navigateur : aucune sauvegarde permanente n'est activée."
    )

if "cg_messages" not in st.session_state:
    st.session_state.cg_messages = []
if "cg_outline" not in st.session_state:
    st.session_state.cg_outline = ""
if "cg_approved_outline" not in st.session_state:
    st.session_state.cg_approved_outline = ""
if "cg_sections" not in st.session_state:
    st.session_state.cg_sections = []
if "cg_generated_section" not in st.session_state:
    st.session_state.cg_generated_section = ""

assistant_tab, writing_tab, correction_tab = st.tabs([
    "💬 Assistant intelligent", "📝 Recherche & rédaction", "✍️ Correction de documents"
])


with correction_tab:
    st.subheader("Correction automatique de fichiers Word")
    st.write(
        "Importez un document .docx pour corriger l’orthographe, la grammaire, "
        "les accords et certaines erreurs de conjugaison. Le moteur en ligne "
        "LanguageTool est nécessaire pour l’analyse complète."
    )
    st.warning(
        "La correction automatique peut modifier une formulation à tort. "
        "Conservez l’original et relisez toujours le document corrigé."
    )
    correction_language = st.selectbox(
        "Langue du document à corriger",
        ["Français", "English"],
        key="cg_correction_language",
    )
    uploaded_docx = st.file_uploader(
        "Importer un document Word (.docx)",
        type=["docx"],
        key="cg_docx_upload",
    )
    if uploaded_docx is not None:
        st.caption(f"Fichier sélectionné : {uploaded_docx.name}")
        if st.button("Lancer la correction automatique", type="primary", key="cg_run_correction"):
            try:
                source_bytes = uploaded_docx.getvalue()
                source_document = Document(BytesIO(source_bytes))
                diagnostics = analyze_document_structure(source_document)
                language_code = "fr" if correction_language == "Français" else "en-US"
                with st.spinner("Analyse et correction du document…"):
                    corrected_document, applied, skipped, detected, suggestions = (
                        correct_docx_with_languagetool(source_document, language_code)
                    )
                output = BytesIO()
                corrected_document.save(output)
                st.session_state.cg_corrected_docx = output.getvalue()
                st.session_state.cg_correction_report = {
                    "filename": uploaded_docx.name,
                    "detected": detected,
                    "applied": applied,
                    "skipped": skipped,
                    "suggestions": suggestions,
                    "diagnostics": diagnostics,
                }
            except Exception as exc:
                st.error(
                    "La correction n’a pas pu aboutir. Vérifiez votre connexion "
                    "Internet ou réessayez plus tard."
                )
                st.caption(f"Détail technique : {type(exc).__name__}")

    report = st.session_state.get("cg_correction_report")
    corrected_bytes = st.session_state.get("cg_corrected_docx")
    if report and corrected_bytes:
        st.divider()
        st.markdown("#### Résultat de la correction")
        c1, c2, c3 = st.columns(3)
        c1.metric("Erreurs / points détectés", report["detected"])
        c2.metric("Corrections appliquées", len(report["applied"]))
        c3.metric("Suggestions non appliquées", len(report["skipped"]))
        st.download_button(
            "Télécharger le Word corrigé (.docx)",
            data=corrected_bytes,
            file_name="correctgloss_corrige.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
        try:
            pdf_bytes = build_pdf_from_docx(corrected_bytes)
            st.download_button(
                "Télécharger le PDF corrigé",
                data=pdf_bytes,
                file_name="correctgloss_corrige.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        except Exception as exc:
            st.info("L’export PDF n’a pas pu être préparé pour ce document.")
            st.caption(f"Détail technique : {type(exc).__name__}")

        with st.expander("Voir les corrections appliquées", expanded=True):
            if report["applied"]:
                st.dataframe(report["applied"], use_container_width=True, hide_index=True)
            else:
                st.info("Aucune correction automatique n’a été appliquée.")
        with st.expander("Voir les suggestions à vérifier"):
            if report["skipped"]:
                st.dataframe(report["skipped"], use_container_width=True, hide_index=True)
            else:
                st.write("Aucune suggestion n’a été laissée en attente.")
        with st.expander("Diagnostic de la structure du document"):
            diagnostics = report["diagnostics"]
            st.write(f"Paragraphes analysés : {diagnostics['paragraph_count']}")
            if diagnostics["repeated_paragraphs"]:
                st.markdown("**Paragraphes répétés à contrôler**")
                st.dataframe(diagnostics["repeated_paragraphs"], use_container_width=True, hide_index=True)
            if diagnostics["suspicious_artefacts"]:
                st.markdown("**Éléments suspects à vérifier manuellement**")
                st.dataframe(diagnostics["suspicious_artefacts"], use_container_width=True, hide_index=True)
            if not diagnostics["repeated_paragraphs"] and not diagnostics["suspicious_artefacts"]:
                st.success("Aucun problème structurel détecté par les contrôles disponibles.")

with assistant_tab:
    st.subheader("Posez votre question")
    st.write(
        "Définition, explication de cours, méthode, exemple ou aide académique. "
        "Vérifiez les informations importantes avant de les utiliser."
    )
    for message in st.session_state.cg_messages:
        with st.chat_message(message["role"]):
            st.markdown(message["content"])

    question = st.chat_input("Écrivez votre question ici…", key="cg_question")
    if question:
        st.session_state.cg_messages.append({"role": "user", "content": question})
        with st.chat_message("user"):
            st.markdown(question)
        with st.chat_message("assistant"):
            with st.spinner("CORRECTGLOSS prépare une réponse…"):
                try:
                    answer = answer_question(
                        question=question,
                        history=st.session_state.cg_messages[:-1],
                        language=language,
                    )
                    st.markdown(answer)
                    st.session_state.cg_messages.append(
                        {"role": "assistant", "content": answer}
                    )
                except (AIConfigurationError, AIProviderError) as exc:
                    st.error(str(exc))
                except Exception as exc:
                    st.error("Une erreur inattendue est survenue. Vérifiez les journaux.")
                    st.caption(type(exc).__name__)

    if st.button("Nouvelle conversation", key="cg_new_chat"):
        st.session_state.cg_messages = []
        st.rerun()

with writing_tab:
    st.subheader("Créer un document académique")
    st.info(
        "Travaillez étape par étape : préparez le plan, validez-le, puis rédigez "
        "chaque partie. Le nombre de pages final dépend de la mise en page et du contenu réel."
    )
    with st.form("cg_outline_form"):
        subject = st.text_input("Sujet du travail", placeholder="Ex. : La compétitivité à l'exportation")
        document_type = st.selectbox(
            "Type de document",
            ["Projet de recherche", "Mémoire", "Rapport de stage", "Dissertation", "Exposé"],
        )
        academic_level = st.selectbox(
            "Niveau académique",
            ["Licence", "Master 1", "Master 2", "Doctorat", "Autre"],
        )
        target_pages = st.number_input(
            "Nombre de pages souhaité", min_value=3, max_value=100, value=20, step=1,
        )
        writing_language = st.selectbox(
            "Langue du document", ["Français", "Malagasy", "English"], index=0,
        )
        instructions = st.text_area(
            "Consignes / établissement / méthodologie (facultatif)", height=90,
        )
        make_outline = st.form_submit_button("1. Générer le plan détaillé", type="primary")

    if make_outline:
        try:
            with st.spinner("Préparation du plan…"):
                st.session_state.cg_outline = create_outline(
                    subject, document_type, academic_level, int(target_pages),
                    writing_language, instructions,
                )
            st.session_state.cg_approved_outline = ""
            st.session_state.cg_sections = []
            st.success("Plan préparé. Relisez-le et corrigez-le avant validation.")
        except (AIConfigurationError, AIProviderError, ValueError) as exc:
            st.error(str(exc))

    if st.session_state.cg_outline:
        st.markdown("#### Plan détaillé à relire")
        edited_outline = st.text_area(
            "Vous pouvez modifier le plan avant de le valider.",
            value=st.session_state.cg_outline,
            height=300,
            key="cg_outline_editor",
        )
        col_a, col_b = st.columns(2)
        with col_a:
            if st.button("2. Valider le plan"):
                st.session_state.cg_approved_outline = edited_outline
                st.success("Plan validé. Vous pouvez rédiger les sections.")
        with col_b:
            if st.button("Réinitialiser le brouillon"):
                st.session_state.cg_outline = ""
                st.session_state.cg_approved_outline = ""
                st.session_state.cg_sections = []
                st.rerun()

    if st.session_state.cg_approved_outline:
        st.divider()
        st.markdown("#### Rédaction progressive")
        section_title = st.text_input(
            "Titre de la section à rédiger",
            placeholder="Ex. : Introduction générale",
        )
        user_material = st.text_area(
            "Vos notes, données, entretiens ou sources (facultatif mais recommandé)",
            height=120,
        )
        target_words = st.slider("Longueur visée pour cette section (mots)", 300, 1800, 1200, 100)
        if st.button("3. Rédiger cette section", type="primary"):
            try:
                previous = "\n\n".join(
                    f"{x['title']}: {x['content'][:700]}" for x in st.session_state.cg_sections[-6:]
                )
                with st.spinner("Rédaction de la section…"):
                    st.session_state.cg_generated_section = write_section(
                        subject=subject,
                        approved_outline=st.session_state.cg_approved_outline,
                        section_title=section_title,
                        user_material=user_material,
                        previous_summary=previous,
                        language=writing_language,
                        target_words=int(target_words),
                    )
            except (AIConfigurationError, AIProviderError, ValueError) as exc:
                st.error(str(exc))

        if st.session_state.cg_generated_section:
            st.markdown("#### Section générée — à vérifier et modifier")
            edited_section = st.text_area(
                "Contenu de la section",
                value=st.session_state.cg_generated_section,
                height=350,
                key="cg_section_editor",
            )
            if st.button("Ajouter cette section au document"):
                if not section_title.strip():
                    st.warning("Donnez un titre à la section.")
                else:
                    st.session_state.cg_sections.append(
                        {"title": section_title.strip(), "content": edited_section}
                    )
                    st.session_state.cg_generated_section = ""
                    st.success("Section ajoutée au document.")
                    st.rerun()

    if st.session_state.cg_sections:
        st.divider()
        st.markdown(f"#### Document en cours ({len(st.session_state.cg_sections)} sections)")
        total_words = 0
        for index, item in enumerate(st.session_state.cg_sections, 1):
            count = len(item["content"].split())
            total_words += count
            with st.expander(f"{index}. {item['title']} — {count} mots"):
                st.write(item["content"])
                if st.button("Supprimer cette section", key=f"cg_delete_{index}"):
                    st.session_state.cg_sections.pop(index - 1)
                    st.rerun()
        st.caption(f"Total actuel : {total_words:,} mots. Le nombre de pages sera vérifié après export.")
        export_title = st.text_input("Titre de la page de garde", value="Document académique")
        author = st.text_input("Nom de l'auteur (facultatif)")
        try:
            docx_bytes = build_docx(
                export_title, subject, writing_language, st.session_state.cg_sections, author,
            )
            pdf_bytes = build_pdf(
                export_title, subject, writing_language, st.session_state.cg_sections, author,
            )
            c1, c2 = st.columns(2)
            with c1:
                st.download_button(
                    "Télécharger Word (.docx)", data=docx_bytes,
                    file_name="correctgloss_document.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )
            with c2:
                st.download_button(
                    "Télécharger PDF", data=pdf_bytes,
                    file_name="correctgloss_document.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
        except Exception as exc:
            st.error("Erreur pendant la préparation des documents exportés.")
            st.caption(type(exc).__name__)

st.divider()
st.caption(
    "Prototype V3 : les réponses et textes générés peuvent contenir des erreurs. "
    "Les citations et références ne sont pas encore vérifiées automatiquement. "
    "N'inventez pas de données de terrain ; contrôlez les sources avant remise académique."
)

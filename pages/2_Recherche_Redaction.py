"""Page de recherche documentaire gratuite de CORRECTGLOSS V2."""
import io
import streamlit as st
from docx import Document
from research_engine import search_web, search_bibliography
from export_engine import build_pdf_from_docx

st.title("🌐 Recherche & rédaction")
st.caption("Recherche gratuite sur le Web et dans des catalogues bibliographiques publics. Aucun abonnement ni clé API payante requis.")
st.info("Les résultats sont des pistes documentaires : ouvrez les notices originales et vérifiez les auteurs, dates, éditions et contenus avant toute citation.")

query = st.text_input(
    "Sujet de recherche",
    placeholder="Ex. : La qualité du girofle malgache et sa compétitivité à l’exportation",
    max_chars=300,
)
c1, c2 = st.columns(2)
with c1:
    max_results = st.selectbox("Nombre de résultats", [5, 8, 10], index=1)
with c2:
    language = st.selectbox("Langue de recherche Web", ["Français", "Anglais"])

left, right = st.columns(2)
with left:
    if st.button("🔎 Rechercher sur Internet", type="primary", use_container_width=True):
        if len(query.strip()) < 3:
            st.warning("Saisissez un sujet d’au moins 3 caractères.")
        else:
            search_query = query.strip() if language == "Français" else query.strip() + " English"
            try:
                with st.spinner("Recherche Web en cours…"):
                    st.session_state["cg_research_results"] = search_web(search_query, max_results=max_results)
                st.session_state["cg_research_query"] = query.strip()
                st.success("Recherche terminée.")
            except Exception as exc:
                st.error(f"Recherche indisponible : {exc}")
with right:
    if st.button("📚 Chercher auteurs, livres et travaux", use_container_width=True):
        if len(query.strip()) < 3:
            st.warning("Saisissez un sujet d’au moins 3 caractères.")
        else:
            try:
                with st.spinner("Consultation des catalogues publics…"):
                    st.session_state["cg_bibliography"] = search_bibliography(query.strip(), max_results=max_results)
                st.session_state["cg_research_query"] = query.strip()
                st.success("Recherche bibliographique terminée.")
            except Exception as exc:
                st.error(f"Recherche bibliographique indisponible : {exc}")

results = st.session_state.get("cg_research_results", [])
if results:
    st.subheader(f"Résultats Web ({len(results)})")
    for i, item in enumerate(results, 1):
        st.markdown(f"**{i}. [{item.get('title', 'Titre indisponible')}]({item.get('url', '')})**")
        if item.get("snippet"):
            st.write(item["snippet"])
        st.caption(f"Pertinence indicative : {round(float(item.get('relevance', 0)) * 100)} %")
    st.warning("Un extrait de moteur de recherche ne remplace pas la lecture de la page originale.")

bibliography = st.session_state.get("cg_bibliography", {})
if bibliography:
    st.divider()
    st.subheader("Auteurs et références bibliographiques")
    st.info(bibliography.get("_message", "Résultats issus de catalogues publics."))
    for diagnostic in bibliography.get("_diagnostics", []):
        st.caption(diagnostic)

    authors = bibliography.get("Auteurs et chercheurs", [])
    with st.expander(f"Auteurs et chercheurs ({len(authors)})", expanded=bool(authors)):
        if not authors:
            st.caption("Aucun auteur identifié dans les notices retenues.")
        for i, item in enumerate(authors, 1):
            st.markdown(f"**{i}. {item.get('name', 'Auteur non renseigné')}**")
            if item.get("publications"):
                for pub in item["publications"]:
                    pub_title = pub.get("title", "Publication sans titre")
                    pub_url = pub.get("url", "")
                    group = pub.get("relevance_group", "C")
                    st.markdown(f"- **[{group}]** [{pub_title}]({pub_url})" if pub_url else f"- **[{group}]** {pub_title}")

    group_details = {
        "A": ("A — Références directes", "Travaux qui mentionnent explicitement le girofle et Madagascar."),
        "B": ("B — Références comparatives internationales", "Travaux sur le girofle dans d'autres pays, utiles pour comparer les marchés et la compétitivité."),
        "C": ("C — Références connexes", "Travaux sur Madagascar et plusieurs thèmes liés à la qualité, l'exportation ou la compétitivité, sans mention directe du girofle."),
    }
    reference_categories = ["Articles scientifiques", "Livres", "Documents, mémoires et thèses"]
    for group in ("A", "B", "C"):
        grouped_items = {
            category: [item for item in bibliography.get(category, [])
                       if item.get("relevance_group", "C") == group]
            for category in reference_categories
        }
        total = sum(len(items) for items in grouped_items.values())
        heading, explanation = group_details[group]
        with st.expander(f"{heading} ({total})", expanded=(group == "A" and total > 0)):
            st.caption(explanation)
            if not total:
                st.caption("Aucun résultat vérifié/pertinent dans cette catégorie pour le moment.")
            for category, entries in grouped_items.items():
                if entries:
                    st.markdown(f"#### {category}")
                for i, item in enumerate(entries, 1):
                    st.markdown(f"**{i}. {item.get('title', 'Référence sans titre')}**")
                    if item.get("authors"):
                        st.caption("Auteur(s) : " + ", ".join(item["authors"]))
                    meta = []
                    if item.get("year"):
                        meta.append(str(item["year"]))
                    if item.get("provider"):
                        meta.append("Catalogue : " + item["provider"])
                    if item.get("doi"):
                        meta.append("DOI : " + str(item["doi"]).replace("https://doi.org/", ""))
                    if meta:
                        st.caption(" • ".join(meta))
                    if item.get("url", "").startswith(("https://", "http://")):
                        st.markdown(f"[Consulter la notice originale]({item['url']})")
                    if item.get("doi_status"):
                        st.caption("État du lien/DOI (ne garantit pas le téléchargement du PDF) : " + str(item["doi_status"]))
                    if item.get("metadata_status"):
                        st.warning(item["metadata_status"])
                    if item.get("relevance_reason"):
                        st.caption(item["relevance_reason"])
                    if item.get("snippet"):
                        st.write(item["snippet"])

    st.warning("Les résultats ne garantissent ni l’exactitude de la notice ni la qualité scientifique du document. Vérifiez chaque source avant de la citer.")

    if st.button("📥 Préparer la bibliographie Word et PDF"):
        try:
            doc = Document()
            doc.add_heading("Bibliographie – CORRECTGLOSS", 0)
            doc.add_paragraph("Sujet : " + st.session_state.get("cg_research_query", query))
            export_groups = {
                "A — Références directes": "A",
                "B — Références comparatives internationales": "B",
                "C — Références connexes à vérifier": "C",
            }
            for group_title, group_code in export_groups.items():
                doc.add_heading(group_title, level=1)
                for category in ["Articles scientifiques", "Livres", "Documents, mémoires et thèses"]:
                    entries = [item for item in bibliography.get(category, [])
                               if item.get("relevance_group", "C") == group_code]
                    if entries:
                        doc.add_heading(category, level=2)
                    for item in entries:
                        p = doc.add_paragraph(style="List Number")
                        p.add_run(item.get("title", "Référence sans titre")).bold = True
                        details = []
                        if item.get("authors"):
                            details.append(", ".join(item["authors"]))
                        if item.get("year"):
                            details.append(str(item["year"]))
                        if item.get("doi"):
                            details.append("DOI : " + str(item["doi"]).replace("https://doi.org/", ""))
                        if item.get("url"):
                            details.append(item["url"])
                        if item.get("metadata_status"):
                            details.append(item["metadata_status"])
                        if details:
                            doc.add_paragraph(" — ".join(details))
            buffer = io.BytesIO()
            doc.save(buffer)
            docx_bytes = buffer.getvalue()
            st.session_state["cg_bibliography_docx"] = docx_bytes
            try:
                st.session_state["cg_bibliography_pdf"] = build_pdf_from_docx(
                    docx_bytes, font="Times New Roman", font_size=12,
                    line_spacing="1.5", margin_cm=2.5,
                )
                st.session_state.pop("cg_bibliography_pdf_error", None)
            except Exception as pdf_exc:
                st.session_state.pop("cg_bibliography_pdf", None)
                st.session_state["cg_bibliography_pdf_error"] = str(pdf_exc)
            st.success("Bibliographie Word préparée.")
        except Exception as exc:
            st.error(f"Erreur de préparation : {exc}")

if st.session_state.get("cg_bibliography_docx"):
    d1, d2 = st.columns(2)
    with d1:
        st.download_button(
            "Télécharger Word", st.session_state["cg_bibliography_docx"],
            file_name="correctgloss_bibliographie.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    with d2:
        if st.session_state.get("cg_bibliography_pdf"):
            st.download_button(
                "Télécharger PDF", st.session_state["cg_bibliography_pdf"],
                file_name="correctgloss_bibliographie.pdf", mime="application/pdf",
            )
        elif st.session_state.get("cg_bibliography_pdf_error"):
            st.caption("Conversion PDF indisponible : " + st.session_state["cg_bibliography_pdf_error"])

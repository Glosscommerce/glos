"""Standalone CORRECTGLOSS V3 Assistant test app.

Run with: streamlit run v3_app.py
This entrypoint is intentionally separate from app.py (the V2 application).
"""
from __future__ import annotations

import os
import streamlit as st

# Read Streamlit secrets into the environment expected by v3_free.ai_provider.
# Secrets remain outside the repository and are never displayed.
try:
    for key in (
        "AI_PROVIDER", "HF_TOKEN", "HF_MODEL", "HF_CHAT_COMPLETIONS_URL",
        "AI_MAX_TOKENS", "OLLAMA_BASE_URL", "OLLAMA_MODEL",
    ):
        value = st.secrets.get(key)
        if value is not None and str(value).strip():
            os.environ[key] = str(value).strip()
except Exception:
    # Local development may run without a Streamlit secrets file.
    pass

from v3_free.ai_provider import (  # noqa: E402
    AIConfigurationError,
    AIProviderError,
    provider_status,
)
from v3_free.assistant import answer_question  # noqa: E402

st.set_page_config(
    page_title="CORRECTGLOSS V3 — Assistant",
    page_icon="📘",
    layout="wide",
)

st.title("📘 CORRECTGLOSS V3")
st.caption("Assistant intelligent — version de test indépendante de V2")

with st.sidebar:
    st.header("⚙️ Configuration")
    language = st.selectbox(
        "Langue de réponse",
        ["Français", "Malagasy", "English"],
        index=0,
    )
    status = provider_status()
    st.markdown("**Moteur IA :** " + status["provider"])
    if status["configured"] == "no":
        st.warning(
            "Le moteur distant n'est pas configuré. Ajoutez HF_TOKEN et HF_MODEL "
            "dans les secrets de cette application de test."
        )
    else:
        st.info(status["note"])
    st.divider()
    if st.button("Nouvelle conversation", use_container_width=True):
        st.session_state.cg_messages = []
        st.rerun()
    st.caption("Les conversations de cette maquette sont conservées dans la session "
               "du navigateur uniquement. Elles ne sont pas encore enregistrées en base.")

if "cg_messages" not in st.session_state:
    st.session_state.cg_messages = []

st.subheader("Posez votre question")
st.write(
    "Vous pouvez demander une définition, une explication de cours, une méthode, "
    "un exemple ou de l'aide pour vos travaux académiques."
)

for message in st.session_state.cg_messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

question = st.chat_input("Écrivez votre question ici…")
if question:
    st.session_state.cg_messages.append({"role": "user", "content": question})
    with st.chat_message("user"):
        st.markdown(question)

    with st.chat_message("assistant"):
        with st.spinner("CORRECTGLOSS prépare une réponse…"):
            try:
                # Pass previous messages only; answer_question appends the current question.
                history = st.session_state.cg_messages[:-1]
                answer = answer_question(
                    question=question,
                    history=history,
                    language=language,
                )
                st.markdown(answer)
                st.session_state.cg_messages.append(
                    {"role": "assistant", "content": answer}
                )
            except AIConfigurationError as exc:
                st.error(str(exc))
                st.info(
                    "Pour tester sans frais API : utilisez Ollama sur un ordinateur "
                    "ou serveur accessible, ou configurez un compte Hugging Face "
                    "et vérifiez les quotas gratuits disponibles."
                )
            except AIProviderError as exc:
                st.error(str(exc))
            except Exception as exc:
                st.error("Une erreur inattendue est survenue. Vérifiez la configuration "
                         "et les journaux de l'application.")
                st.caption(type(exc).__name__)

st.divider()
st.caption(
    "Prototype V3 : les réponses IA peuvent contenir des erreurs. Vérifiez les faits, "
    "les références et les citations avant toute utilisation académique."
)

"""Conversation assistant for CORRECTGLOSS V3."""
from __future__ import annotations
from typing import Dict, List
from .ai_provider import generate_text

SYSTEM_PROMPTS = {
    "Français": (
        "Tu es CORRECTGLOSS, un assistant pédagogique et académique. Réponds clairement, "
        "adapte l'explication au niveau demandé et donne un exemple si utile. Distingue "
        "faits vérifiés, hypothèses et incertitudes. N'invente jamais source, citation, "
        "statistique ou référence bibliographique."
    ),
    "Malagasy": (
        "Ianao no CORRECTGLOSS, mpanampy amin'ny fianarana sy asa akademika. Valio amin'ny "
        "teny Malagasy mazava sy mora azo. Omeo ohatra raha ilaina. Saraho ny zava-misy "
        "voamarina sy ny vinavina. Aza mamorona loharano, citation na statistika."
    ),
    "English": (
        "You are CORRECTGLOSS, an educational and academic assistant. Answer clearly, "
        "adapt to the requested level, and give examples when useful. Distinguish verified "
        "facts from assumptions. Never invent sources, citations or statistics."
    ),
}

def answer_question(
    question: str, history: List[Dict[str, str]] | None = None,
    language: str = "Français",
) -> str:
    question = (question or "").strip()
    if not question:
        raise ValueError("Veuillez saisir une question.")
    messages: List[Dict[str, str]] = [{
        "role": "system",
        "content": SYSTEM_PROMPTS.get(language, SYSTEM_PROMPTS["Français"]),
    }]
    for item in (history or [])[-8:]:
        role, content = item.get("role"), item.get("content", "")
        if role in ("user", "assistant") and isinstance(content, str) and content.strip():
            messages.append({"role": role, "content": content[-4000:]})
    messages.append({"role": "user", "content": question})
    return generate_text(messages)

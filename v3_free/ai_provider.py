"""Free-first AI provider adapter for CORRECTGLOSS V3.

Routes:
1) Ollama on a machine/server you control (no API key; requires a running model).
2) Hugging Face Inference Providers using a user-supplied account token.
Free quotas and model availability are controlled by the provider and can change.
"""
from __future__ import annotations
import os
from typing import Dict, List, Optional
import requests

class AIConfigurationError(RuntimeError):
    """The free AI engine is not configured."""

class AIProviderError(RuntimeError):
    """The configured AI provider returned an error."""

def _ollama_chat(messages: List[Dict[str, str]], model: str, timeout: int) -> str:
    base_url = os.getenv("OLLAMA_BASE_URL", "http://127.0.0.1:11434").rstrip("/")
    payload = {
        "model": model or os.getenv("OLLAMA_MODEL", "qwen2.5:3b"),
        "messages": messages, "stream": False,
        "options": {"temperature": 0.4},
    }
    try:
        response = requests.post(f"{base_url}/api/chat", json=payload, timeout=timeout)
        response.raise_for_status()
        answer = response.json().get("message", {}).get("content", "").strip()
        if not answer:
            raise AIProviderError("Ollama n'a renvoyé aucun texte.")
        return answer
    except requests.RequestException as exc:
        raise AIProviderError(
            "Impossible de joindre Ollama. Vérifiez que le serveur est démarré "
            "et que le modèle est installé."
        ) from exc

def _huggingface_chat(messages: List[Dict[str, str]], model: str, timeout: int) -> str:
    token = os.getenv("HF_TOKEN", "").strip()
    if not token:
        raise AIConfigurationError(
            "Aucun moteur IA distant n'est configuré. Ajoutez HF_TOKEN dans les "
            "secrets Streamlit, ou configurez Ollama sur votre serveur."
        )
    selected_model = model or os.getenv("HF_MODEL", "Qwen/Qwen2.5-7B-Instruct")
    endpoint = os.getenv(
        "HF_CHAT_COMPLETIONS_URL",
        "https://router.huggingface.co/v1/chat/completions",
    ).rstrip("/")
    payload = {
        "model": selected_model, "messages": messages,
        "temperature": 0.4,
        "max_tokens": int(os.getenv("AI_MAX_TOKENS", "1200")),
    }
    try:
        response = requests.post(
            endpoint,
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json=payload, timeout=timeout,
        )
        if response.status_code in (401, 403):
            raise AIProviderError("Jeton Hugging Face invalide ou non autorisé.")
        if response.status_code == 429:
            raise AIProviderError(
                "Quota gratuit atteint ou capacité temporairement indisponible. "
                "Réessayez plus tard."
            )
        response.raise_for_status()
        answer = response.json().get("choices", [{}])[0].get("message", {}).get("content", "")
        if not isinstance(answer, str) or not answer.strip():
            raise AIProviderError("Le fournisseur IA n'a renvoyé aucun texte.")
        return answer.strip()
    except requests.Timeout as exc:
        raise AIProviderError("Le moteur IA a dépassé le délai d'attente.") from exc
    except requests.RequestException as exc:
        raise AIProviderError(f"Erreur du fournisseur IA : {exc}") from exc

def generate_text(
    messages: List[Dict[str, str]], *, provider: Optional[str] = None,
    model: Optional[str] = None, timeout: int = 90,
) -> str:
    """Generate text using a configured free route; never silently fake a response."""
    selected = (provider or os.getenv("AI_PROVIDER", "huggingface")).strip().lower()
    if selected == "ollama":
        return _ollama_chat(messages, model or "", timeout)
    if selected in ("huggingface", "hf"):
        return _huggingface_chat(messages, model or "", timeout)
    raise AIConfigurationError("AI_PROVIDER doit être 'huggingface' ou 'ollama'.")

def provider_status() -> Dict[str, str]:
    """Return safe configuration status without exposing credentials."""
    provider = os.getenv("AI_PROVIDER", "huggingface").strip().lower()
    if provider == "ollama":
        return {
            "provider": "Ollama",
            "configured": "yes" if os.getenv("OLLAMA_BASE_URL") else "local default",
            "note": "Le modèle doit être installé et le serveur accessible.",
        }
    return {
        "provider": "Hugging Face",
        "configured": "yes" if os.getenv("HF_TOKEN") else "no",
        "note": "Le quota gratuit dépend du compte et du modèle.",
    }

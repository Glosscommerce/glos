"""Recherche web pour CORRECTGLOSS.

Les résultats sont fournis par DuckDuckGo via le paquet ddgs.
Les extraits sont des pistes de recherche : l'utilisateur doit ouvrir et vérifier
chaque source avant de citer les informations dans un travail académique.
"""
from ddgs import DDGS


def search_web(query: str, max_results: int = 8):
    query = (query or "").strip()
    if len(query) < 3:
        raise ValueError("Saisissez un sujet de recherche d'au moins 3 caractères.")
    if len(query) > 300:
        raise ValueError("Le sujet est trop long (300 caractères maximum).")

    results = DDGS().text(
        query,
        region="wt-wt",
        safesearch="moderate",
        max_results=max(1, min(int(max_results), 10)),
    )

    cleaned = []
    seen = set()
    for item in results or []:
        url = (item.get("href") or "").strip()
        title = (item.get("title") or "").strip()
        body = (item.get("body") or "").strip()
        if not url.startswith(("https://", "http://")) or url in seen:
            continue
        seen.add(url)
        cleaned.append({
            "title": title[:300] or "Source sans titre",
            "url": url,
            "snippet": body[:1200],
        })
    return cleaned

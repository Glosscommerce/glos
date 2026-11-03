"""Recherche académique gratuite pour CORRECTGLOSS.

Sources publiques : OpenAlex, Crossref, Open Library et DDGS.
Aucune clé OpenAI ni API payante n'est nécessaire. Les résultats sont des pistes
bibliographiques : toujours vérifier la notice originale avant citation.
"""
import re
import unicodedata
from urllib.parse import urlparse

from ddgs import DDGS


STOP_WORDS = {
    "le", "la", "les", "un", "une", "des", "du", "de", "et", "en", "a", "au",
    "aux", "pour", "sur", "dans", "par", "avec", "the", "and", "of", "in",
    "to", "for", "on", "with", "from", "an", "une", "d", "l"
}
LOW_QUALITY_DOMAINS = {
    "wordhippo.com", "reverso.net", "definitions.net", "grokipedia.com"
}
ACADEMIC_DOMAINS = (
    "doi.org", "openalex.org", "crossref.org", "openlibrary.org",
    "scholar.google.", "sciencedirect.com", "springer.com", "wiley.com",
    "tandfonline.com", "researchgate.net", "jstor.org", "theses.fr",
    "hal.science", "core.ac.uk", "oatd.org", "univ-", ".edu", ".ac."
)


def _normalise(value):
    value = unicodedata.normalize("NFKD", str(value or "").lower())
    value = "".join(c for c in value if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]+", " ", value).strip()


def _terms(value):
    return {word for word in _normalise(value).split()
            if len(word) > 2 and word not in STOP_WORDS}


def _relevance(topic, title="", snippet="", url=""):
    terms = _terms(topic)
    if not terms:
        return 0.0
    title_terms = _terms(title)
    body_terms = _terms(snippet)
    title_score = len(terms & title_terms) / len(terms)
    body_score = len(terms & body_terms) / len(terms)
    domain = urlparse(url).netloc.lower()
    academic_bonus = 0.15 if any(d in domain for d in ACADEMIC_DOMAINS) else 0
    return min(1.0, title_score * 0.70 + body_score * 0.30 + academic_bonus)


def search_web(query: str, max_results: int = 8):
    query = (query or "").strip()
    if len(query) < 3:
        raise ValueError("Saisissez un sujet de recherche d'au moins 3 caractères.")
    if len(query) > 300:
        raise ValueError("Le sujet est trop long (300 caractères maximum).")

    raw = DDGS().text(query, region="wt-wt", safesearch="moderate",
                      max_results=max(10, min(int(max_results) * 3, 30)))
    cleaned, seen = [], set()
    for item in raw or []:
        url = (item.get("href") or "").strip()
        title = (item.get("title") or "").strip()
        body = (item.get("body") or "").strip()
        if not url.startswith(("https://", "http://")) or url in seen or not title:
            continue
        domain = urlparse(url).netloc.lower().removeprefix("www.")
        if any(domain == bad or domain.endswith("." + bad) for bad in LOW_QUALITY_DOMAINS):
            continue
        score = _relevance(query, title, body, url)
        if score < 0.18:
            continue
        seen.add(url)
        cleaned.append({
            "title": title[:300], "url": url, "snippet": body[:1200],
            "relevance": round(score, 2),
        })
    cleaned.sort(key=lambda item: item["relevance"], reverse=True)
    return cleaned[:max(1, min(int(max_results), 10))]


def search_bibliography(topic: str, max_results: int = 8):
    """Recherche des références et conserve leurs métadonnées vérifiables."""
    import requests

    topic = (topic or "").strip()
    if len(topic) < 3:
        raise ValueError("Saisissez un thème de recherche d'au moins 3 caractères.")
    limit = max(1, min(int(max_results), 10))
    headers = {"User-Agent": "CORRECTGLOSS/1.1 (academic reference discovery)"}
    output = {
        "Articles scientifiques": [], "Livres": [],
        "Auteurs et chercheurs": [], "Documents, mémoires et thèses": [],
    }
    seen = set()
    errors = []

    def add_reference(category, title, year=None, authors=None, url=None,
                      provider="", doi=None, abstract=None, work_type=None):
        title = re.sub(r"\s+", " ", str(title or "")).strip()
        if not title or not url:
            return
        score = _relevance(topic, title, abstract or "", url)
        # Les catalogues renvoient parfois des résultats éloignés : on les écarte.
        if score < 0.12:
            return
        key = (category, _normalise(title))
        if key in seen:
            return
        seen.add(key)
        author_list = [str(a).strip() for a in (authors or []) if str(a).strip()][:6]
        details = []
        if doi:
            details.append("DOI : " + str(doi).replace("https://doi.org/", ""))
        if work_type:
            details.append("Type : " + str(work_type))
        if abstract:
            details.append("Résumé : " + re.sub(r"\s+", " ", str(abstract))[:650])
        output[category].append({
            "title": title[:300], "year": year, "authors": author_list,
            "url": url, "type": work_type or "publication",
            "provider": provider, "relevance": round(score, 2), "doi": doi,
            "snippet": " | ".join(details),
        })
        for author in author_list:
            author_key = _normalise(author)
            if author_key and not any(_normalise(x["name"]) == author_key
                                      for x in output["Auteurs et chercheurs"]):
                output["Auteurs et chercheurs"].append({
                    "name": author, "url": url, "provider": provider,
                })

    # OpenAlex : catalogue ouvert de publications et de chercheurs.
    try:
        response = requests.get(
            "https://api.openalex.org/works",
            params={"search": topic, "per-page": min(limit * 3, 30),
                    "select": "id,doi,display_name,publication_year,type,authorships,primary_location"},
            headers=headers, timeout=18,
        )
        response.raise_for_status()
        for work in response.json().get("results", []):
            title = work.get("display_name") or ""
            authors = [(a.get("author") or {}).get("display_name")
                       for a in work.get("authorships", [])[:6]]
            location = work.get("primary_location") or {}
            url = location.get("landing_page_url") or work.get("doi") or work.get("id")
            kind = (work.get("type") or "").lower()
            category = "Livres" if kind in {"book", "monograph"} else "Articles scientifiques"
            add_reference(category, title, work.get("publication_year"), authors,
                          url, "OpenAlex", work.get("doi"), work_type=kind)
    except (requests.RequestException, ValueError, TypeError) as exc:
        errors.append("OpenAlex : " + str(exc)[:180])

    # Crossref : DOI, auteurs et date de publication.
    try:
        response = requests.get(
            "https://api.crossref.org/works",
            params={"query.bibliographic": topic, "rows": min(limit * 3, 30)},
            headers=headers, timeout=18,
        )
        response.raise_for_status()
        for item in response.json().get("message", {}).get("items", []):
            authors = [" ".join(filter(None, [a.get("given"), a.get("family")]))
                       for a in item.get("author", [])[:6]]
            date_parts = (item.get("published", {}).get("date-parts") or [[None]])[0]
            doi = item.get("DOI")
            url = ("https://doi.org/" + doi) if doi else item.get("URL")
            add_reference("Articles scientifiques",
                          (item.get("title") or [""])[0],
                          date_parts[0] if date_parts else None, authors, url,
                          "Crossref", doi, work_type=item.get("type"))
    except (requests.RequestException, ValueError, TypeError, IndexError) as exc:
        errors.append("Crossref : " + str(exc)[:180])

    # Open Library : livres et notices de catalogue.
    try:
        response = requests.get(
            "https://openlibrary.org/search.json",
            params={"q": topic, "limit": min(limit * 3, 30),
                    "fields": "title,author_name,first_publish_year,key"},
            headers=headers, timeout=18,
        )
        response.raise_for_status()
        for item in response.json().get("docs", []):
            key = item.get("key") or ""
            if not key.startswith("/works/"):
                continue
            add_reference("Livres", item.get("title"),
                          item.get("first_publish_year"), item.get("author_name", []),
                          "https://openlibrary.org" + key, "Open Library",
                          work_type="book")
    except (requests.RequestException, ValueError, TypeError) as exc:
        errors.append("Open Library : " + str(exc)[:180])

    # Recherche de documents : requête ciblée et filtrage de pertinence.
    try:
        doc_query = f'{topic} mémoire thèse dissertation rapport PDF'
        raw_docs = DDGS().text(doc_query, region="wt-wt", safesearch="moderate",
                               max_results=min(limit * 3, 30))
        for item in raw_docs or []:
            url = (item.get("href") or "").strip()
            title = (item.get("title") or "").strip()
            snippet = (item.get("body") or "").strip()
            if not url.startswith(("https://", "http://")):
                continue
            score = _relevance(topic, title, snippet, url)
            if score < 0.22:
                continue
            output["Documents, mémoires et thèses"].append({
                "title": title[:300], "url": url, "snippet": snippet[:700],
                "provider": "Recherche web", "relevance": round(score, 2),
            })
    except Exception as exc:
        errors.append("Recherche de documents : " + str(exc)[:180])

    for category in ("Articles scientifiques", "Livres", "Documents, mémoires et thèses"):
        output[category].sort(key=lambda x: x.get("relevance", 0), reverse=True)
        output[category] = output[category][:limit]
    output["Auteurs et chercheurs"] = output["Auteurs et chercheurs"][:limit * 2]
    output["_diagnostics"] = errors
    output["_message"] = (
        "Certaines bases n'ont pas répondu : " + "; ".join(errors)
        if errors else "Recherche terminée. Vérifiez chaque notice auprès de sa source."
    )
    return output

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


def search_bibliography(topic: str, max_results: int = 8):
    """Trouve des références bibliographiques via OpenAlex, Crossref et Open Library."""
    import requests

    topic = (topic or "").strip()
    if len(topic) < 3:
        raise ValueError("Saisissez un thème d'au moins 3 caractères.")
    limit = max(1, min(int(max_results), 10))
    headers = {"User-Agent": "CORRECTGLOSS/1.0 (academic reference discovery)"}
    output = {"Articles scientifiques": [], "Livres": [], "Auteurs et chercheurs": []}
    seen = set()

    # OpenAlex : articles, chapitres, livres académiques et auteurs associés.
    try:
        response = requests.get(
            "https://api.openalex.org/works",
            params={"search": topic, "per-page": limit, "select":
                    "id,display_name,publication_year,type,doi,authorships,primary_location"},
            headers=headers, timeout=20,
        )
        response.raise_for_status()
        for work in response.json().get("results", []):
            title = (work.get("display_name") or "").strip()
            if not title:
                continue
            authors = []
            for authorship in work.get("authorships", [])[:5]:
                author = (authorship.get("author") or {}).get("display_name")
                if author and author not in authors:
                    authors.append(author)
            location = work.get("primary_location") or {}
            source = location.get("landing_page_url") or work.get("doi") or work.get("id")
            kind = (work.get("type") or "").lower()
            category = "Livres" if kind in {"book", "monograph"} else "Articles scientifiques"
            key = (category, title.lower())
            if key not in seen:
                seen.add(key)
                output[category].append({
                    "title": title, "year": work.get("publication_year"),
                    "authors": authors, "url": source, "type": work.get("type", "publication"),
                    "provider": "OpenAlex",
                })
            for author in authors:
                if author.lower() not in {x["name"].lower() for x in output["Auteurs et chercheurs"]}:
                    output["Auteurs et chercheurs"].append({
                        "name": author, "url": work.get("id"), "provider": "OpenAlex"
                    })
    except requests.RequestException:
        pass

    # Crossref : références de revues et publications avec métadonnées.
    try:
        response = requests.get(
            "https://api.crossref.org/works",
            params={"query.bibliographic": topic, "rows": limit},
            headers=headers, timeout=20,
        )
        response.raise_for_status()
        for item in response.json().get("message", {}).get("items", []):
            title = ((item.get("title") or [""])[0] or "").strip()
            if not title:
                continue
            authors = [
                " ".join(filter(None, [a.get("given"), a.get("family")])).strip()
                for a in item.get("author", [])[:5]
            ]
            authors = [a for a in authors if a]
            url = item.get("URL") or item.get("DOI")
            key = ("Articles scientifiques", title.lower())
            if key not in seen:
                seen.add(key)
                output["Articles scientifiques"].append({
                    "title": title, "year": (item.get("published", {}).get("date-parts") or [[None]])[0][0],
                    "authors": authors, "url": url, "type": item.get("type", "publication"),
                    "provider": "Crossref",
                })
            for author in authors:
                if author.lower() not in {x["name"].lower() for x in output["Auteurs et chercheurs"]}:
                    output["Auteurs et chercheurs"].append({
                        "name": author, "url": url, "provider": "Crossref"
                    })
    except (requests.RequestException, ValueError, TypeError, IndexError):
        pass

    # Open Library : catalogues de livres, avec lien vers la notice.
    try:
        response = requests.get(
            "https://openlibrary.org/search.json",
            params={"q": topic, "limit": limit, "fields": "title,author_name,first_publish_year,key"},
            headers=headers, timeout=20,
        )
        response.raise_for_status()
        for item in response.json().get("docs", []):
            title = (item.get("title") or "").strip()
            if not title:
                continue
            authors = item.get("author_name", [])[:5]
            key = ("Livres", title.lower())
            if key not in seen:
                seen.add(key)
                output["Livres"].append({
                    "title": title, "year": item.get("first_publish_year"),
                    "authors": authors, "url": "https://openlibrary.org" + item.get("key", ""),
                    "type": "book", "provider": "Open Library",
                })
            for author in authors:
                if author.lower() not in {x["name"].lower() for x in output["Auteurs et chercheurs"]}:
                    output["Auteurs et chercheurs"].append({
                        "name": author, "url": "https://openlibrary.org" + item.get("key", ""),
                        "provider": "Open Library"
                    })
    except requests.RequestException:
        pass

    # Recherche complémentaire de mémoires, thèses, rapports et documents PDF.
    documents = []
    try:
        doc_query = f'{topic} (mémoire OR thesis OR dissertation OR report OR rapport OR filetype:pdf)'
        for item in DDGS().text(doc_query, region="wt-wt", safesearch="moderate", max_results=limit):
            url = (item.get("href") or "").strip()
            title = (item.get("title") or "").strip()
            if url.startswith(("https://", "http://")) and title:
                documents.append({
                    "title": title[:300], "url": url,
                    "snippet": (item.get("body") or "")[:700],
                    "provider": "Recherche web",
                })
    except Exception:
        pass
    output["Documents, mémoires et thèses"] = documents[:limit]

    for category in ("Articles scientifiques", "Livres"):
        output[category] = output[category][:limit]
    output["Auteurs et chercheurs"] = output["Auteurs et chercheurs"][:limit * 2]
    return output

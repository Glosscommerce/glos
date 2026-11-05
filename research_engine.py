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
    """Score thématique tolérant aux formulations différentes, sans survaloriser le domaine."""
    terms = _terms(topic)
    if not terms:
        return 0.0
    title_terms = _terms(title)
    body_terms = _terms(snippet)
    # Le titre compte davantage que le résumé ; les mots-clés présents dans les deux
    # renforcent le score. Cela évite qu'un résultat partageant un seul mot générique
    # soit classé comme très pertinent.
    title_score = len(terms & title_terms) / len(terms)
    body_score = len(terms & body_terms) / len(terms)
    overlap = len(terms & (title_terms | body_terms)) / len(terms)
    domain = urlparse(url).netloc.lower()
    academic_bonus = 0.06 if any(d in domain for d in ACADEMIC_DOMAINS) else 0
    # Bonus de contexte, sans considérer un simple mot générique comme preuve suffisante.
    context_bonus = 0.04 if any(x in (title_terms | body_terms)
                                for x in {"madagascar", "malgache", "girofle", "clove"}) else 0
    return min(1.0, title_score * 0.55 + body_score * 0.20
               + overlap * 0.25 + academic_bonus + context_bonus)


def _verify_reference_link(url, doi=None):
    """Vérifie la résolution du DOI ou l'accessibilité HTTP de la notice."""
    import requests
    target = ("https://doi.org/" + str(doi).replace("https://doi.org/", "").strip()
              if doi else str(url or "").strip())
    if not target.startswith(("https://", "http://")):
        return "non vérifié"
    try:
        response = requests.head(
            target, allow_redirects=True,
            headers={"User-Agent": "CORRECTGLOSS/1.2 (reference link check)"},
            timeout=7,
        )
        if response.status_code < 400:
            return "accessible"
        # Certains dépôts refusent HEAD mais acceptent GET.
        response = requests.get(
            target, allow_redirects=True, stream=True,
            headers={"User-Agent": "CORRECTGLOSS/1.2 (reference link check)"},
            timeout=7,
        )
        return "accessible" if response.status_code < 400 else "à vérifier"
    except requests.RequestException:
        return "à vérifier"


def _openalex_abstract(inverted_index):
    """Reconstruit le résumé OpenAlex à partir de son index de positions."""
    if not isinstance(inverted_index, dict):
        return ""
    words = []
    for word, positions in inverted_index.items():
        for position in positions or []:
            if isinstance(position, int) and 0 <= position < 1200:
                words.append((position, word))
    return " ".join(word for _, word in sorted(words))[:5000]


def _topic_queries(topic):
    """Variantes courtes pour récupérer des publications utilisant d'autres termes."""
    variants = [topic]
    terms = list(_terms(topic))
    if any(x in terms for x in {"girofle", "clou", "cloves", "clove"}):
        variants.extend([
            "clove quality Madagascar export competitiveness",
            "Madagascar clove value chain quality export",
            "filière girofle Madagascar qualité exportation"
        ])
    return list(dict.fromkeys(variants))


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
        title_body = _terms(title + " " + (abstract or ""))
        topic_terms = _terms(topic)
        # Pour un sujet centré sur le girofle, éviter les mémoires sur des filières
        # sans lien identifiable avec le girofle/clove.
        clove_topic = bool(topic_terms & {"girofle", "clove", "cloves"})
        clove_words = {"girofle", "clove", "cloves", "syzygium", "aromaticum"}
        clove_match = bool(title_body & clove_words)
        topic_norm = _normalise(topic)
        title_norm = _normalise(title)
        # Si le sujet porte explicitement sur Madagascar, ne pas confondre
        # les études indonésiennes ou d'autres pays avec le cas malgache.
        madagascar_topic = "madagascar" in topic_norm or "malgache" in topic_norm
        madagascar_match = any(term in (title_norm + " " + _normalise(abstract or ""))
                               for term in ("madagascar", "malgache", "malagasy"))
        # Écarter les travaux consacrés à une autre filière (litchi, vanille,
        # raphia...) même s'ils parlent aussi d'exportation et de qualité.
        other_crop_words = {
            "litchi", "lychee", "vanille", "vanilla", "raphia",
            "cacao", "cocoa", "cafe", "coffee", "poivre", "pepper"
        }
        other_crop_only = bool(title_body & other_crop_words) and not bool(
            _terms(title) & clove_words
        )
        # Écarter les faux positifs où « clous de girofle » est une expression
        # figurée sans rapport avec la filière agricole (ex. gaullisme).
        figurative_terms = {"gaullisme", "gaulliste"}
        figurative_match = bool(title_body & figurative_terms)
        if clove_topic and (other_crop_only or not clove_match or figurative_match):
            return
        if clove_topic and madagascar_topic and not madagascar_match:
            return
        # Pour un mémoire qualité/exportation, exiger au moins un axe
        # directement visible dans le titre ou le résumé.
        topic_axes = topic_terms & {
            "qualite", "quality", "export", "exportation", "competitivite",
            "competitiveness", "marche", "market", "filiere", "value", "chain"
        }
        if clove_topic and topic_axes:
            axis_terms = _terms(title + " " + (abstract or ""))
            if not (topic_axes & axis_terms) and score < 0.35:
                return
        # Seuil relevé pour limiter les résultats faiblement liés au sujet.
        if score < 0.20:
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
        doi_status = _verify_reference_link(url, doi) if doi else "absent"
        source_status = _verify_reference_link(url) if not doi else doi_status
        reference = {
            "title": title[:300], "year": year, "authors": author_list,
            "url": url, "type": work_type or "publication",
            "provider": provider, "relevance": round(score, 2), "doi": doi,
            "doi_status": doi_status, "source_status": source_status,
            "snippet": " | ".join(details),
        }
        output[category].append(reference)
        for author in author_list:
            author_key = _normalise(author)
            if not author_key:
                continue
            author_entry = next(
                (x for x in output["Auteurs et chercheurs"]
                 if _normalise(x["name"]) == author_key), None
            )
            if author_entry is None:
                author_entry = {
                    "name": author, "url": url, "provider": provider,
                    "publications": [],
                }
                output["Auteurs et chercheurs"].append(author_entry)
            if not any(_normalise(pub.get("title")) == _normalise(title)
                       for pub in author_entry["publications"]):
                author_entry["publications"].append({
                    "title": reference["title"], "year": year,
                    "category": category, "url": url, "doi": doi,
                    "provider": provider, "type": work_type or "publication",
                })

    # Requêtes courtes complémentaires : le titre d'un livre n'emploie pas toujours
    # les mêmes mots que le sujet exact du mémoire.
    queries = _topic_queries(topic)
    # OpenAlex : catalogue ouvert de publications et de chercheurs.
    for search_term in queries[:2]:
        try:
            response = requests.get(
                "https://api.openalex.org/works",
                params={"search": search_term, "per-page": min(limit * 2, 20),
                        "select": "id,doi,display_name,publication_year,type,authorships,primary_location,abstract_inverted_index"},
                headers=headers, timeout=18,
            )
            if response.status_code == 429:
                errors.append("OpenAlex : limite temporaire de requêtes (HTTP 429). Les autres bases continuent.")
                break
            response.raise_for_status()
            for work in response.json().get("results", []):
                title = work.get("display_name") or ""
                authors = [(a.get("author") or {}).get("display_name")
                           for a in work.get("authorships", [])[:6]]
                location = work.get("primary_location") or {}
                url = location.get("landing_page_url") or work.get("doi") or work.get("id")
                kind = (work.get("type") or "").lower()
                if kind in {"dissertation", "thesis"}:
                    category = "Documents, mémoires et thèses"
                elif kind in {"book", "monograph", "book-chapter", "edited-book"}:
                    category = "Livres"
                elif kind in {"article", "review", "preprint"}:
                    category = "Articles scientifiques"
                else:
                    continue
                abstract = _openalex_abstract(work.get("abstract_inverted_index"))
                add_reference(category, title, work.get("publication_year"), authors,
                              url, "OpenAlex", work.get("doi"), abstract=abstract,
                              work_type=kind)
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
            crossref_type = (item.get("type") or "").lower()
            crossref_title = (item.get("title") or [""])[0]
            if crossref_type in {"dissertation", "thesis"}:
                category = "Documents, mémoires et thèses"
            elif crossref_type in {"journal-article", "proceedings-article", "posted-content"}:
                category = "Articles scientifiques"
            elif crossref_type in {"book", "monograph", "edited-book", "book-chapter"}:
                category = "Livres"
            else:
                # Ne pas présenter les composants, rapports et notices diverses
                # comme des articles scientifiques évalués par les pairs.
                continue
            add_reference(category, crossref_title,
                          date_parts[0] if date_parts else None, authors, url,
                          "Crossref", doi, work_type=crossref_type)
    except (requests.RequestException, ValueError, TypeError, IndexError) as exc:
        errors.append("Crossref : " + str(exc)[:180])

    # Google Books : une recherche par sujet principal et une par variante ciblée.
    # En cas de quota 429, on arrête immédiatement pour éviter des appels inutiles.
    google_queries = queries[:2]
    for search_term in google_queries:
        try:
            response = requests.get(
                "https://www.googleapis.com/books/v1/volumes",
                params={"q": search_term, "maxResults": min(limit, 8),
                        "orderBy": "relevance", "printType": "books"},
                headers=headers, timeout=18,
            )
            if response.status_code == 429:
                break
            response.raise_for_status()
            for volume in response.json().get("items", []):
                info = volume.get("volumeInfo") or {}
                access = volume.get("accessInfo") or {}
                volume_id = volume.get("id")
                link = info.get("previewLink") or info.get("infoLink")
                if not link and volume_id:
                    link = "https://books.google.com/books?id=" + volume_id
                before_count = len(output["Livres"])
                add_reference(
                    "Livres", info.get("title"),
                    (info.get("publishedDate") or "")[:4] or None,
                    info.get("authors", []), link, "Google Books",
                    abstract=info.get("description"), work_type="book",
                )
                if len(output["Livres"]) > before_count:
                    output["Livres"][-1]["google_books_id"] = volume_id
                    output["Livres"][-1]["viewability"] = access.get("viewability", "NO_PAGES")
                    output["Livres"][-1]["preview_link"] = info.get("previewLink", "")
        except (requests.RequestException, ValueError, TypeError) as exc:
            errors.append("Google Books (facultatif) : " + str(exc)[:180])
            break

    # Open Library : interroger plusieurs formulations, puis dédoublonner les notices.
    for search_term in queries[:3]:
        try:
            response = requests.get(
                "https://openlibrary.org/search.json",
                params={"q": search_term, "limit": min(limit * 2, 20),
                        "fields": "title,author_name,first_publish_year,key,subject,ia,has_fulltext,public_scan,edition_key,cover_i"},
                headers=headers, timeout=18,
            )
            response.raise_for_status()
            for item in response.json().get("docs", []):
                key = item.get("key") or ""
                if not key.startswith("/works/"):
                    continue
                subjects = ", ".join((item.get("subject") or [])[:12])
                before_count = len(output["Livres"])
                add_reference("Livres", item.get("title"),
                              item.get("first_publish_year"), item.get("author_name", []),
                              "https://openlibrary.org" + key, "Open Library",
                              abstract=subjects, work_type="book")
                if len(output["Livres"]) > before_count:
                    book = output["Livres"][-1]
                    book["open_library_key"] = key
                    book["archive_ids"] = item.get("ia") or []
                    book["has_fulltext"] = bool(item.get("has_fulltext"))
                    book["public_scan"] = bool(item.get("public_scan"))
                    book["edition_keys"] = item.get("edition_key") or []
                    book["cover_id"] = item.get("cover_i")
        except (requests.RequestException, ValueError, TypeError) as exc:
            errors.append("Open Library : " + str(exc)[:180])
            break

    # Les mémoires et thèses sont déjà classés dans la recherche OpenAlex générale.
    # Éviter un appel OpenAlex supplémentaire qui augmente le risque de HTTP 429.

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

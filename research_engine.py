"""Recherche académique gratuite pour CORRECTGLOSS.

Sources publiques : OpenAlex, Crossref, Open Library et DDGS.
Aucune clé OpenAI ni API payante n'est nécessaire. Les résultats sont des pistes
bibliographiques : toujours vérifier la notice originale avant citation.
"""
import re
import unicodedata
import copy
import threading
import time
import requests
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


def _check_link_details(url, doi=None):
    """Vérifie un lien sans télécharger le document ; distingue page et PDF."""
    target = ("https://doi.org/" + str(doi).replace("https://doi.org/", "").strip()
              if doi else str(url or "").strip())
    result = {
        "status": "non vérifié",
        "http_status": None,
        "final_url": target,
        "content_type": "",
        "is_pdf": False,
        "detail": "URL absente ou non prise en charge.",
    }
    if not target.startswith(("https://", "http://")):
        return result

    response = None
    try:
        # GET en streaming est plus compatible que HEAD avec les dépôts académiques.
        # Le corps n'est jamais téléchargé : on lit uniquement les en-têtes.
        response = requests.get(
            target, allow_redirects=True, stream=True,
            headers={"User-Agent": "Mozilla/5.0 (compatible; CORRECTGLOSS link checker/2.1)"},
            timeout=(5, 8),
        )
        status = response.status_code
        final_url = response.url
        content_type = (response.headers.get("Content-Type") or "").split(";")[0].strip().lower()
        is_pdf = (
            "application/pdf" in content_type
            or urlparse(final_url).path.lower().endswith(".pdf")
        )
        result.update({
            "http_status": status,
            "final_url": final_url,
            "content_type": content_type,
            "is_pdf": is_pdf,
        })
        if 200 <= status < 300 and is_pdf:
            result.update(status="pdf", detail="Le serveur répond et indique un fichier PDF. La gratuité et la licence restent à vérifier.")
        elif 200 <= status < 300:
            result.update(status="page", detail="La page de destination répond. Le téléchargement d'un PDF n'est pas confirmé.")
        elif status in (401, 403, 405, 429):
            result.update(status="blocked", detail="Le serveur bloque ou limite la vérification automatique ; essayez d'ouvrir le lien dans votre navigateur.")
        elif status in (404, 410):
            result.update(status="dead", detail="La page ou le document semble introuvable (HTTP {}).".format(status))
        elif status >= 500:
            result.update(status="server_error", detail="Le serveur distant rencontre une erreur (HTTP {}).".format(status))
        else:
            result.update(status="uncertain", detail="Réponse HTTP {} : disponibilité à confirmer.".format(status))
    except requests.RequestException as exc:
        result.update(status="uncertain", detail="Vérification impossible depuis le serveur CORRECTGLOSS : " + str(exc)[:140])
    finally:
        if response is not None:
            response.close()
    return result


def _verify_reference_link(url, doi=None):
    """Compatibilité : renvoie un libellé court pour les anciennes interfaces."""
    details = _check_link_details(url, doi)
    return {
        "pdf": "Le serveur indique un PDF — ouverture dans le navigateur non garantie",
        "page": "le serveur répond — ouverture dans le navigateur non garantie",
        "blocked": "accès bloqué au vérificateur",
        "dead": "lien probablement indisponible",
        "server_error": "erreur du serveur distant",
        "uncertain": "à vérifier dans le navigateur",
    }.get(details["status"], "non vérifié")


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
    # Ordre déterministe : _terms() renvoie un set et ne doit pas décider
    # aléatoirement des requêtes envoyées aux catalogues.
    terms = set(_normalise(topic).split())
    if terms & {"girofle", "clou", "cloves", "clove"}:
        variants.extend([
            "clove quality Madagascar export competitiveness",
            "Madagascar clove value chain quality export",
            "filière girofle Madagascar qualité exportation",
            "Madagascar agricultural product quality export competitiveness",
            "Madagascar agricultural value chains export market competitiveness"
        ])
    return list(dict.fromkeys(variants))


def search_web(query: str, max_results: int = 8):
    query = (query or "").strip()
    if len(query) < 3:
        raise ValueError("Saisissez un sujet de recherche d'au moins 3 caractères.")
    if len(query) > 300:
        raise ValueError("Le sujet est trop long (300 caractères maximum).")

    # Plusieurs moteurs de recherche sont essayés : une panne DuckDuckGo
    # ne doit pas interrompre toute la recherche Internet.
    raw = []
    backend_errors = []
    requested_results = max(10, min(int(max_results) * 3, 30))
    for backend in ("duckduckgo", "bing", "brave"):
        try:
            raw = list(DDGS(timeout=8).text(
                query,
                region="wt-wt",
                safesearch="moderate",
                max_results=requested_results,
                backend=backend,
            ) or [])
            if raw:
                break
        except Exception as exc:
            backend_errors.append(f"{backend}: {str(exc)[:160]}")
    if not raw and backend_errors:
        raise RuntimeError(
            "Les moteurs de recherche sont temporairement indisponibles. "
            + " | ".join(backend_errors)
        )
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
    selected = cleaned[:max(1, min(int(max_results), 10))]
    for item in selected:
        link_check = _check_link_details(item["url"])
        item.update({
            "link_status": link_check["status"],
            "link_label": {
                "pdf": "Le serveur indique un PDF — ouverture dans le navigateur non garantie",
                "page": "Le serveur répond — ouverture dans le navigateur non garantie",
                "blocked": "Vérification bloquée — ouvrir dans le navigateur",
                "dead": "Lien probablement indisponible",
                "server_error": "Erreur du serveur distant",
                "uncertain": "À vérifier dans le navigateur",
            }.get(link_check["status"], "Non vérifié"),
            "link_detail": link_check["detail"],
            "final_url": link_check["final_url"],
            "content_type": link_check["content_type"],
        })
    return selected


# Cache mémoire court : évite de répéter les mêmes appels lors des reruns Streamlit.
_BIBLIO_CACHE = {}
_BIBLIO_CACHE_LOCK = threading.Lock()
_BIBLIO_CACHE_TTL = 6 * 60 * 60  # 6 heures
_BIBLIO_CACHE_MAX = 64


def search_bibliography(topic: str, max_results: int = 8):
    """Recherche avec cache mémoire et protection contre les limites temporaires."""
    import requests
    from requests.adapters import HTTPAdapter
    from urllib3.util.retry import Retry

    topic = (topic or "").strip()
    limit = max(1, min(int(max_results), 10))
    cache_key = (_normalise(topic), limit)
    now = time.time()
    with _BIBLIO_CACHE_LOCK:
        cached = _BIBLIO_CACHE.get(cache_key)
        if cached and now - cached["time"] < _BIBLIO_CACHE_TTL:
            return copy.deepcopy(cached["data"])
        if cached:
            _BIBLIO_CACHE.pop(cache_key, None)

    retry_policy = Retry(
        total=2, connect=2, read=2, status=2,
        backoff_factor=1.5,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset(["GET"]),
        respect_retry_after_header=True,
        raise_on_status=False,
    )
    session = requests.Session()
    session.mount("https://", HTTPAdapter(max_retries=retry_policy))
    session.mount("http://", HTTPAdapter(max_retries=retry_policy))

    headers = {"User-Agent": "CORRECTGLOSS/2.0 (academic bibliography search)"}

    if len(topic) < 3:
        raise ValueError("Saisissez un thème de recherche d'au moins 3 caractères.")
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
        title_terms = _terms(title)
        abstract_terms = _terms(abstract or "")
        title_body = title_terms | abstract_terms
        topic_terms = _terms(topic)
        topic_norm = _normalise(topic)
        title_norm = _normalise(title)
        clove_topic = bool(topic_terms & {"girofle", "clove", "cloves"})
        clove_words = {"girofle", "clove", "cloves", "syzygium", "aromaticum"}
        clove_match = bool(title_body & clove_words)
        abstract_clove_match = bool(abstract_terms & clove_words)
        madagascar_match = any(
            term in (title_norm + " " + _normalise(abstract or ""))
            for term in ("madagascar", "malgache", "malagasy")
        )
        madagascar_title_match = any(
            term in title_norm for term in ("madagascar", "malgache", "malagasy")
        )
        other_crop_words = {
            "litchi", "lychee", "vanille", "vanilla", "raphia",
            "cacao", "cocoa", "cafe", "coffee", "poivre", "pepper",
            "cranberry", "crab", "citrus", "cinnamon"
        }
        other_crop_only = bool(title_body & other_crop_words) and not bool(
            title_terms & clove_words
        )
        figurative_match = bool(title_body & {"gaullisme", "gaulliste"})
        relevance_group = "A"
        relevance_reason = "Correspondance directe avec le sujet."

        # Traitement strict des sujets culturels centrés sur le Tsaboraha :
        # « malagasy » seul ne constitue jamais une preuve de pertinence directe.
        tsaboraha_topic = bool(topic_terms & {"tsaboraha", "tsboraha", "tsaborage"})
        betsimisaraka_topic = "betsimisaraka" in topic_terms
        if tsaboraha_topic:
            tsaboraha_words = {"tsaboraha", "tsboraha", "tsaborage"}
            betsimisaraka_match = "betsimisaraka" in title_body
            tsaboraha_match = bool(title_body & tsaboraha_words)
            if tsaboraha_match and betsimisaraka_match:
                relevance_group = "A"
                relevance_reason = "La notice mentionne explicitement le Tsaboraha et les Betsimisaraka."
            elif tsaboraha_match:
                relevance_group = "B"
                relevance_reason = "La notice mentionne le Tsaboraha, mais le lien spécifique avec les Betsimisaraka reste à confirmer."
            else:
                cultural_context = {
                    "rituel", "rituels", "ritual", "rituals", "coutume", "coutumes",
                    "custom", "customs", "culture", "cultural", "tradition",
                    "traditions", "ancetre", "ancetres", "ancestor", "ancestors",
                    "tromba", "fanokoarana", "famadihana", "funeraire", "funeral",
                    "mort", "death", "religion", "spirit", "esprit"
                }
                context_match = bool(title_body & cultural_context)
                if betsimisaraka_match and context_match:
                    relevance_group = "C"
                    relevance_reason = "Référence connexe sur la culture, les rites ou les traditions Betsimisaraka ; elle ne traite pas explicitement du Tsaboraha."
                else:
                    return
        elif clove_topic:
            if figurative_match:
                return
            if clove_match:
                if other_crop_only:
                    return
                if madagascar_title_match:
                    relevance_group = "A"
                    relevance_reason = "Le titre porte explicitement sur le girofle à Madagascar."
                else:
                    relevance_group = "B"
                    relevance_reason = "Étude sur le girofle à l'international ; utile pour comparaison."
            else:
                # Écarter les notices dont le titre porte clairement sur une autre
                # culture (litchi, vanille, cacao...), même si elles parlent aussi
                # de qualité, d'exportation ou de Madagascar.
                title_other_crop = bool(title_terms & other_crop_words) and not bool(
                    title_terms & clove_words
                )
                if title_other_crop:
                    return
                topic_axes = {
                    "qualite", "quality", "export", "exportation", "exportations",
                    "competitivite", "competitiveness", "marche", "market",
                    "filiere", "value", "chain", "commerce", "trade"
                }
                axis_hits = title_body & topic_axes
                if not madagascar_match or len(axis_hits) < 2:
                    return
                relevance_group = "C"
                relevance_reason = "Source connexe sur Madagascar et plusieurs axes qualité/export/compétitivité, sans mention directe du girofle."
        else:
            # Pour les autres sujets, exiger une correspondance avec les mots
            # distinctifs du thème ; les termes génériques (importance, étude,
            # analyse, contribution, malagasy...) ne suffisent pas.
            generic_terms = {
                "importance", "impact", "effet", "effets", "etude", "analyse",
                "contribution", "amelioration", "developpement", "role", "place",
                "cas", "exemple", "approche", "problematique", "malagasy",
                "madagascar", "malgache", "chez", "dans", "pour"
            }
            anchors = topic_terms - generic_terms
            anchor_hits = anchors & title_body
            if not anchors or not anchor_hits:
                return
            if betsimisaraka_topic and "betsimisaraka" not in title_body:
                return
            if len(anchors) >= 2 and len(anchor_hits) < 1:
                return
            # Classement générique pour tous les domaines :
            # A = tous les termes distinctifs sont retrouvés ;
            # B = plusieurs termes distinctifs concordent ;
            # C = un seul indice pertinent, à utiliser comme contexte seulement.
            if len(anchor_hits) == len(anchors) and score >= 0.18:
                relevance_group = "A"
                relevance_reason = "Tous les termes distinctifs du sujet sont retrouvés dans la notice ; vérifier le contenu original."
            elif len(anchor_hits) >= 2:
                relevance_group = "B"
                relevance_reason = "Plusieurs termes distinctifs concordent ; source comparative à vérifier."
            else:
                relevance_group = "C"
                relevance_reason = "Correspondance partielle ; source de contexte uniquement, à vérifier."

        minimum_score = 0.08 if (
            (clove_topic and relevance_group in {"A", "B"}) or tsaboraha_topic
        ) else 0.12
        if score < minimum_score:
            return
        # Détecter les notices dont le titre et le résumé semblent parler de sujets différents.
        abstract_crop_mismatch = (
            clove_match
            and bool(abstract_terms & other_crop_words)
            and not abstract_clove_match
        )
        distinctive_topic_terms = topic_terms - {
            "importance", "impact", "etude", "analyse", "malagasy", "madagascar"
        }
        abstract_has_topic = bool(abstract_terms & distinctive_topic_terms)
        # Reconnaître les équivalents multilingues girofle/clove pour éviter
        # les fausses alertes lorsque le titre est en français et le résumé en anglais.
        if clove_topic and abstract_terms & clove_words:
            abstract_has_topic = True
        abstract_topic_mismatch = bool(
            abstract_terms and distinctive_topic_terms
            and (title_terms & distinctive_topic_terms)
            and not abstract_has_topic
        )
        metadata_status = (
            "ALERTE : le résumé semble concerner une autre filière que le titre ; ne pas citer avant vérification."
            if abstract_crop_mismatch else
            "ALERTE : les mots distinctifs du thème apparaissent dans le titre mais pas dans le résumé ; vérifier la notice originale."
            if abstract_topic_mismatch else
            "Notice bibliographique non validée automatiquement : le lien peut être accessible sans que le titre, les auteurs ou le contenu soient confirmés."
        )
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
        link_check = _check_link_details(url, doi)
        short_link_status = {
            "pdf": "Le serveur indique un PDF — ouverture dans le navigateur non garantie",
            "page": "le serveur répond — ouverture dans le navigateur non garantie",
            "blocked": "accès bloqué au vérificateur",
            "dead": "lien probablement indisponible",
            "server_error": "erreur du serveur distant",
            "uncertain": "à vérifier dans le navigateur",
        }.get(link_check["status"], "non vérifié")
        doi_status = short_link_status if doi else "absent"
        source_status = short_link_status
        reference = {
            "title": title[:300], "year": year, "authors": author_list,
            "url": url, "type": work_type or "publication",
            "provider": provider, "relevance": round(score, 2), "doi": doi,
            "doi_status": doi_status, "source_status": source_status,
            "link_status": link_check["status"],
            "link_label": {
                "pdf": "Le serveur indique un PDF — ouverture dans le navigateur non garantie",
                "page": "Le serveur répond — ouverture dans le navigateur non garantie",
                "blocked": "Vérification bloquée — ouvrir dans le navigateur",
                "dead": "Lien probablement indisponible",
                "server_error": "Erreur du serveur distant",
                "uncertain": "À vérifier dans le navigateur",
            }.get(link_check["status"], "Non vérifié"),
            "link_detail": link_check["detail"],
            "final_url": link_check["final_url"],
            "content_type": link_check["content_type"],
            "relevance_group": relevance_group,
            "relevance_reason": relevance_reason,
            "metadata_status": metadata_status,
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
                    "relevance_group": relevance_group,
                })

    # Requêtes courtes complémentaires : le titre d'un livre n'emploie pas toujours
    # les mêmes mots que le sujet exact du mémoire.
    queries = _topic_queries(topic)
    # OpenAlex : catalogue ouvert de publications et de chercheurs.
    for search_term in queries[:5]:
        try:
            response = session.get(
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
        response = session.get(
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
            response = session.get(
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
            response = session.get(
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

    group_order = {"A": 0, "B": 1, "C": 2}
    for category in ("Articles scientifiques", "Livres", "Documents, mémoires et thèses"):
        output[category].sort(
            key=lambda x: (group_order.get(x.get("relevance_group", "C"), 9),
                           -x.get("relevance", 0))
        )
        output[category] = output[category][:limit]
    output["Auteurs et chercheurs"] = output["Auteurs et chercheurs"][:limit * 2]
    output["_diagnostics"] = errors
    output["_message"] = (
        "Certaines bases n'ont pas répondu : " + "; ".join(errors)
        if errors else "Recherche terminée. Vérifiez chaque notice auprès de sa source."
    )
    # Ne jamais mettre en cache une réponse dégradée par une limite API :
    # une nouvelle recherche pourra retenter OpenAlex au lieu de répéter l'ancien 429.
    has_temporary_provider_error = any(
        "HTTP 429" in str(error) or "HTTP 5" in str(error)
        for error in errors
    )
    if not has_temporary_provider_error:
        with _BIBLIO_CACHE_LOCK:
            if len(_BIBLIO_CACHE) >= _BIBLIO_CACHE_MAX:
                oldest_key = min(_BIBLIO_CACHE, key=lambda key: _BIBLIO_CACHE[key]["time"])
                _BIBLIO_CACHE.pop(oldest_key, None)
            _BIBLIO_CACHE[cache_key] = {"time": time.time(), "data": copy.deepcopy(output)}
    return output
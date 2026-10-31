"""Moteur de correction CORRECTGLOSS : suggestions vérifiables et application contrôlée."""
import re
import requests


RULES = [
    (r"\bmalgré que\b", "bien que", "Grammaire", "« Malgré que » est généralement déconseillé dans la rédaction soignée."),
    (r"\bau jour d'aujourd'hui\b", "aujourd'hui", "Style", "Expression redondante."),
    (r"\btrès important\b", "essentiel", "Style", "Choisissez un terme plus précis selon le contexte."),
    (r"\bon peut voir que\b", "les résultats montrent que", "Style académique", "Évitez les formulations vagues."),
    (r"\bje pense que\b", "les données suggèrent que", "Style académique", "Évitez une opinion personnelle non justifiée."),
    (r"\bafin de\s+de\b", "afin de", "Grammaire", "Répétition de la préposition."),
    (r"\bpar contre\b", "en revanche", "Style", "« En revanche » convient souvent mieux au registre académique."),
    (r"\bà cause de\b", "en raison de", "Style", "Formulation plus neutre dans un texte académique."),
    (r"\bceci dit\b", "cela dit", "Style", "Forme usuelle à privilégier."),
    (r"\bmalgrès\b", "malgré", "Orthographe", "Attention à l'accent : « malgré » ne prend pas de s."),
    (r"\b dévelopement\b", "développement", "Orthographe", "Orthographe du mot."),
    (r"\bcommerciauxes\b", "commerciales", "Orthographe", "Accord et forme du mot à vérifier."),
]


def offline_suggestions(content, language="Français"):
    """Retourne des suggestions avec positions pour permettre une application ciblée."""
    if language == "Anglais":
        rules = [
            (r"\bmore better\b", "better", "Grammaire", "Double comparatif incorrect."),
            (r"\binformations\b", "information", "Grammaire", "« Information » est généralement indénombrable en anglais."),
            (r"\bhe don't\b", "he doesn't", "Grammaire", "Accord sujet-verbe."),
            (r"\bshe don't\b", "she doesn't", "Grammaire", "Accord sujet-verbe."),
            (r"\bi has\b", "I have", "Grammaire", "Conjugaison du verbe have."),
        ]
    elif language == "Malagasy":
        rules = []
    else:
        rules = RULES
    found = []
    for pattern, suggestion, category, note in rules:
        for match in re.finditer(pattern, content, re.IGNORECASE):
            found.append({
                "motif": match.group(0), "suggestion": suggestion,
                "category": category, "explication": note,
                "start": match.start(), "end": match.end(),
                "source": "CORRECTGLOSS",
            })
    return found


def languagetool_suggestions(content, language_code):
    """Interroge LanguageTool et convertit ses résultats au format CORRECTGLOSS."""
    response = requests.post(
        "https://api.languagetool.org/v2/check",
        data={"text": content[:20000], "language": language_code},
        timeout=30,
    )
    response.raise_for_status()
    found = []
    for item in response.json().get("matches", []):
        start = int(item.get("offset", 0))
        length = int(item.get("length", 0))
        if length <= 0:
            continue
        replacements = [x.get("value", "") for x in item.get("replacements", [])[:3]]
        found.append({
            "motif": content[start:start + length],
            "suggestion": replacements[0] if replacements else "",
            "alternatives": replacements[1:],
            "category": item.get("rule", {}).get("category", {}).get("name", "Grammaire"),
            "explication": item.get("message", "Suggestion LanguageTool."),
            "start": start, "end": start + length,
            "source": "LanguageTool",
        })
    return found


def apply_suggestions(content, suggestions):
    """Applique les suggestions de droite à gauche pour préserver les positions."""
    usable = [
        item for item in suggestions
        if item.get("suggestion")
        and isinstance(item.get("start"), int)
        and isinstance(item.get("end"), int)
        and 0 <= item["start"] < item["end"] <= len(content)
        and content[item["start"]:item["end"]].casefold() == item.get("motif", "").casefold()
    ]
    # En cas de chevauchement, on conserve une seule suggestion par zone.
    usable.sort(key=lambda x: (x["start"], x["end"]))
    selected = []
    last_end = -1
    for item in usable:
        if item["start"] >= last_end:
            selected.append(item)
            last_end = item["end"]
    for item in reversed(selected):
        content = content[:item["start"]] + item["suggestion"] + content[item["end"]:]
    return content

"""Chapter-by-chapter academic writing helpers for CORRECTGLOSS V3."""
from __future__ import annotations
from .ai_provider import generate_text

def create_outline(
    subject: str, document_type: str, academic_level: str, target_pages: int,
    language: str = "Français", instructions: str = "",
) -> str:
    if not subject.strip():
        raise ValueError("Le sujet est obligatoire.")
    if not 3 <= int(target_pages) <= 100:
        raise ValueError("Le nombre de pages doit être compris entre 3 et 100.")
    prompt = f"""
Prépare un plan détaillé pour un document académique.
Type : {document_type}
Sujet : {subject}
Niveau : {academic_level}
Longueur visée : environ {target_pages} pages
Langue : {language}
Consignes : {instructions or "Aucune"}

Exigences :
- Structure cohérente avec titres et sous-titres numérotés.
- Répartition indicative des pages.
- Problématique, objectifs et hypothèses si pertinents.
- Données et sources nécessaires pour chaque partie.
- N'invente aucun résultat de terrain, chiffre, citation ou référence.
- Marque les éléments qui exigent des informations de l'utilisateur.
- Termine par les questions à valider avant rédaction.
"""
    return generate_text([
        {"role": "system", "content": (
            "Tu es un assistant de méthodologie académique. Propose des plans réalistes "
            "sans fabriquer de données ni de références."
        )},
        {"role": "user", "content": prompt},
    ])

def write_section(
    subject: str, approved_outline: str, section_title: str,
    user_material: str = "", previous_summary: str = "",
    language: str = "Français", target_words: int = 1200,
) -> str:
    if not subject.strip() or not approved_outline.strip() or not section_title.strip():
        raise ValueError("Le sujet, le plan validé et le titre de section sont obligatoires.")
    target_words = max(300, min(int(target_words), 1800))
    prompt = f"""
Rédige uniquement cette section académique.
Sujet : {subject}
Langue : {language}
Section : {section_title}
Plan validé :
{approved_outline[:10000]}

Matériaux fournis par l'utilisateur :
{user_material[:12000] or "Aucun matériau fourni."}

Résumé des sections précédentes :
{previous_summary[:3000] or "Aucune section précédente."}

Consignes :
- Vise environ {target_words} mots sans remplissage artificiel.
- Respecte le plan et évite les répétitions.
- N'invente aucun entretien, résultat, chiffre, norme, citation ou source.
- Si une donnée manque, insère [DONNÉE À FOURNIR].
- N'invente pas de bibliographie; marque [RÉFÉRENCE À VÉRIFIER] si nécessaire.
- Retourne la section avec ses sous-titres, sans préambule conversationnel.
"""
    return generate_text([
        {"role": "system", "content": (
            "Tu es un rédacteur académique prudent. La fiabilité prime sur la longueur."
        )},
        {"role": "user", "content": prompt},
    ])

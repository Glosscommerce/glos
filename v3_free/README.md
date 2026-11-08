# CORRECTGLOSS V3 — premier socle IA gratuit

Ce dossier est isolé de l'application V2. Il ne remplace ni ne modifie app.py.

## Moteurs gratuits envisagés

- Hugging Face Inference Providers : jeton de compte et quota gratuit éventuel. Le quota, les modèles disponibles et les conditions peuvent changer.
- Ollama : modèle open-source exécuté sur un ordinateur ou serveur contrôlé par le projet. Pas de clé API, mais nécessite des ressources matérielles et un serveur accessible.
- Aucun secret ne doit être enregistré dans GitHub.

## Configuration Hugging Face

Installer requests, puis définir AI_PROVIDER=huggingface, HF_TOKEN et HF_MODEL dans l'environnement ou les secrets Streamlit.

Modèle initial proposé : Qwen/Qwen2.5-7B-Instruct. Sa disponibilité réelle doit être testée avec le compte et le fournisseur sélectionné.

## Configuration Ollama

Installer Ollama, télécharger un modèle adapté (par exemple qwen2.5:3b), démarrer le serveur, puis définir AI_PROVIDER=ollama et OLLAMA_MODEL=qwen2.5:3b.

Attention : Ollama sur 127.0.0.1 de l'ordinateur du développeur n'est pas accessible depuis Streamlit Cloud.

## Fonctions livrées dans cette première brique

- answer_question() : assistant avec historique borné.
- create_outline() : plan détaillé.
- write_section() : rédaction progressive, une section à la fois.

Le nombre exact de pages n'est pas encore garanti. Il faudra ajouter l'assemblage DOCX, la mesure réelle des pages, la sauvegarde/reprise, l'interface utilisateur et des tests de cohérence avant de déclarer la rédaction 20–30+ pages prête.
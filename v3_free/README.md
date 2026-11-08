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

## Interface Assistant (prototype)

Le point d'entrée indépendant de V2 est `v3_app.py`.

### Lancer en local

```bash
pip install -r requirements.txt
streamlit run v3_app.py
```

### Déployer une application de test séparée sur Streamlit Community Cloud

1. Créer une nouvelle application Streamlit, sans modifier l'application V2.
2. Repository : `Glosscommerce/glos`.
3. Branch : `correctgloss-v3-free-test`.
4. Main file path : `v3_app.py`.
5. Ajouter les secrets dans les paramètres de cette nouvelle application, par exemple :

```toml
AI_PROVIDER = "huggingface"
HF_TOKEN = "votre_jeton_personnel"
HF_MODEL = "Qwen/Qwen2.5-7B-Instruct"
```

6. Déployer et vérifier les journaux. Ne jamais committer le jeton dans GitHub.

Le modèle configuré doit être effectivement pris en charge par le fournisseur et le compte. Les offres gratuites peuvent être limitées, suspendues ou modifiées. L'interface ne simule pas une réponse si le moteur IA est indisponible.

### Limites connues

- Cette interface est un prototype de chat, pas encore la version finale.
- L'historique est conservé dans la session Streamlit et disparaît selon le cycle de session.
- Aucun document n'est encore généré ou exporté depuis cette interface.
- Les fonctions de rédaction longue existent séparément, mais l'interface de planification et de rédaction chapitre par chapitre reste à construire.

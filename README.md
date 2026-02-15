# ExcelChat

Transformez votre expérience Excel grâce à l'IA. Utilisez un assistant conversationnel pour analyser vos données, manipuler des tableaux et créer des graphiques en langage naturel.

# Installation
```bash
git clone https://github.com/Bsh54/ExcelChat.git
cd ExcelChat
pip install -r requirments.txt
python -m src.main
```

# Fonctionnalités Clés
+ 🤖 **Assistant Conversationnel** : L'IA explique ses actions en français et fournit la logique Excel équivalente.
+ 📄 **Nouvelle Feuille** : Commencez un projet à partir d'un tableau vierge et laissez l'IA générer des données.
+ 📊 **Visualisation** : Création automatique de graphiques avec Matplotlib via de simples commandes.
+ 🔄 **Auto-Correction** : L'agent détecte et corrige seul les erreurs de code Python.
+ 📥 **Exportation** : Exportez facilement vos résultats calculés en .xlsx ou .csv.

# Raccourcis clavier
+ **Ctrl + Q** : Envoyer la question à l'IA
+ **Ctrl + A** : Réduire/Agrandir la fenêtre de chat
+ **Ctrl + Z** : Annuler la modification

# Principes
Utilise **pandas** pour lire les données Excel et représenter les feuilles de calcul sous forme d'objets DataFrame.
Utilise l'IA pour générer le code de traitement du DataFrame, exécute ce code via un interpréteur isolé et renvoie les résultats sur l'interface graphique (GUI).

# Problèmes connus
+ Les fichiers Excel avec des cellules fusionnées ne sont pas supportés.
+ Les en-têtes du fichier Excel doivent impérativement se trouver sur la première ligne.

# Version 23 — 10 septembre 2026

Le conseiller peut reprendre une analyse complète dans une nouvelle session et préparer un rapport énergétique sans configurer de projet solaire.

- **Dossier complet JSON** : fichiers Enedis/GEREDIS et PMA, entreprise, période, horaires HC, tarifs, hypothèses photovoltaïques et financières, statut et annotations. Le profil PVGIS déjà obtenu est également conservé ; il est réutilisé tant que les paramètres solaires restent identiques. Télécharger une nouvelle copie après chaque modification à conserver.
- **Compatibilité** : les brouillons JSON de version 1 restent importables. Ils ne contiennent que les informations et annotations initialement sauvegardées ; les fichiers et réglages existants ne sont pas remplacés lors de leur import.
- **Deux parcours** : l’analyse énergétique fonctionne sans géocodage ni PVGIS. Les paramètres solaires et annotations sont conservés lors des changements de parcours. L’export Excel énergie ne contient pas de simulation financière photovoltaïque fictive.
- **Rapports** : les annotations du PDF photovoltaïque sont réparties dans les sections concernées, y compris les points à approfondir et le statut du dossier. Les deux PDF disposent d’un aperçu de leur mise en page exacte, fondé sur les mêmes fichiers que les téléchargements. Le navigateur doit prendre en charge l’affichage intégré des PDF ; le téléchargement reste disponible.
- **Tarification corrigée** : la synthèse applique le type de tarif sélectionné (unique, HP/HC ou hiver/été). La part fixe annuelle est proratisée sur la période couverte, avec une base de 365,25 jours, cohérente avec l’annualisation financière existante. Les totaux figurent aussi dans la synthèse Excel.
- **Import des dossiers** : validation des versions, champs, types, bornes et intégrité des fichiers avant restauration. Limite de 50 Mo par fichier source et 150 Mo par dossier JSON. Aucun fichier client n’est enregistré dans le dépôt ou sur un serveur de stockage par cette fonctionnalité.

## Vérification

Installer `requirements-dev.txt`, puis exécuter `python -m pytest`.

Les tests utilisent des données fictives : sauvegarde/restauration, anciens brouillons, rejets des dossiers invalides, parcours énergie sans réseau, trois modes tarifaires, PMA seul, changements de parcours et reprise dans une nouvelle session, réutilisation du profil solaire, placement des annotations dans les PDF et export Excel. Les réponses PVGIS sont simulées ; les tests ne valident pas la disponibilité du service externe.

Les fonctions de sauvegarde sont isolées dans `dossiers.py` (format et validation), `dossier_schema.py` (réglages persistants) et `dossier_ui.py` (intégration Streamlit). Les formules de consommation, PMA et Autocalsol ne sont pas modifiées.

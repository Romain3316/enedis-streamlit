# Version 25 — 10 septembre 2026

- Suppression du menu « Espace de travail » : l’analyse s’ouvre directement, les informations et paramètres du dossier restent dans le panneau latéral.
- Le bouton « Préparer le rapport » ouvre les exports et annotations ; « Retour à l’analyse » permet de revenir aux résultats.
- Le choix exclusif « Parcours » devient « Ajouter une étude photovoltaïque ». Les onglets énergie restent disponibles lorsque le complément solaire est activé.
- Compatibilité des dossiers existants et conservation des paramètres solaires et annotations lors des changements d’affichage.

# Version 24 — 10 septembre 2026

- Nouvelle organisation en trois pages : Dossier, Analyse et Rapport. Sauvegarde du dossier et préparation du rapport accessibles en haut de page.
- Interface allégée, logo CMA original et charte bleu marine/rouge conservés. Indicateurs de synthèse regroupés et paramètres accessibles depuis le dossier.
- Profils : matrice complète des 24 heures, courbe semaine/week-end et annotations à proximité. Les cartes détaillées restent disponibles dans un volet dépliable.
- Palette historique vert/jaune/rouge partagée entre tableaux, cartes thermiques, PDF et exports de matrice. Les PDF énergie et photovoltaïque colorent désormais leurs cellules avec les mêmes valeurs non arrondies.
- Conservation des dossiers V23 et des notes lors des changements de page. Vérification automatisée du parcours Analyse → Rapport → Dossier, en complément des tests de restauration, tarifs et photovoltaïque.

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

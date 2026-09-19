# Journal des changements

## 20-09-2026

- 01:00 Ajout d’une installation Windows par utilisateur dans LocalAppData, avec raccourci du menu Démarrer, icône native, enregistrement dans les Applications installées et désinstallation complète sans droits administrateur.
- 01:00 Validation du cycle installé : lancement depuis le dossier système utilisateur, extraction PDF de 8 740 caractères, suppression complète de l’exécutable, du raccourci et du registre, puis réinstallation réussie de Zen 0.1.0.
- 00:52 Préparation de la publication publique sur GitHub : exclusion de la distribution autonome du suivi Git, ajout du lien vers les Releases dans le README et contrôle local de l’absence de secrets.
- 00:50 Enrichissement du README avec la stack C#/.NET 8 et WinUI 3, les bibliothèques documentaires, OCR et images, le fonctionnement local autonome et la commande de publication Windows x64.
- 00:48 Remplacement du README technique par une présentation courte et professionnelle de Zen, avec fonctionnalités essentielles, liens de lancement et capture propre de l’interface Windows 11 actuelle.
- 00:45 Nettoyage du projet : suppression des anciens prototypes Tauri, React et WinForms, des dépendances Node/Rust, des caches de compilation .NET, des publications intermédiaires, des captures de contrôle et du banc de test temporaire, pour un gain total de 802,08 Mo avec les fichiers de test de Téléchargements.
- 00:45 Conservation exclusive du code source WinUI 3, de la distribution autonome Zen-Windows11 et de la documentation active ; remplacement du README Tauri obsolète et adaptation du fichier .gitignore au projet .NET actuel.
- 00:41 Extension de la conversion vers JPG aux images PNG, BMP, GIF et TIFF avec un outil unique « Image en JPG » et des filtres de sélection cohérents.
- 00:41 Remplacement des sorties TXT par une zone de texte native intégrée à Zen pour les extractions PDF, DOCX et OCR, avec bouton « Copier tout » et état de résultat contextuel.
- 00:41 Séparation Clean Code du catalogue des outils et du service d’extraction de texte afin de conserver une fenêtre principale concise et des responsabilités isolées.
- 00:41 Ajout dans Téléchargements d’une photographie CC0 et de deux affiches du domaine public pour tester les conversions d’images et l’OCR en conditions réelles.
- 00:41 Ajout d’un banc de tests UI Automation reproductible et validation par clics des conversions PNG/BMP/GIF/TIFF vers JPG, PNG/JPG vers PDF, des trois extractions, du presse-papiers, du rognage, des coins arrondis et de l’image circulaire.
- 00:41 Contrôle technique des sorties finales : JPG 1920 × 1440 lisibles, PDF avec signature valide, rognage 800 × 800, transparence réelle des coins et découpe circulaire 1440 × 1440.
- 00:11 Intégration de l’icône officielle du projet Zen dans l’exécutable, la barre de titre et le dossier de publication Windows.
- 00:11 Chargement local et temporaire de la famille Exo officielle pendant l’import PDF, sans installation système, puis incorporation de la police dans le DOCX.
- 00:11 Restauration des frontières de pages PDF, y compris les coupures situées au milieu d’un paragraphe, sans modifier les zones de texte positionnées.
- 00:11 Reconstruction des blocs parallèles de signatures en grille Word éditable à deux colonnes afin de conserver leur répartition et leur séparateur sur la page 7.
- 00:11 Conservation de la signature manuscrite sur une page 8 dédiée et suppression des sauts de page ou traits importés en double.
- 00:11 Utilisation prioritaire de Microsoft Word pour l’export DOCX vers PDF fidèle, avec LibreOffice conservé comme solution de secours.
- 00:11 Validation réelle depuis l’interface de test-3.pdf vers DOCX puis du DOCX vers PDF : huit pages A4 dans les deux résultats, police Exo incorporée et aucun processus Word persistant.

## 19-09-2026

- 08:02 Création de Zen, application Windows locale de conversion, extraction et traitement d’images fondée sur Tauri, React et Rust.
- 08:02 Ajout des parcours DOCX/PDF, PDF/DOCX, conversions d’images, génération de PDF, extraction de texte, OCR, rognage, coins arrondis et image circulaire.
- 08:15 Remplacement du socle de compilation par une application Windows WinForms autonome afin de ne nécessiter ni Visual Studio ni autre installation de développement.
- 08:15 Refonte complète de l’interface avec les contrôles et la mise en page adaptative d’une application Windows native : menus, barre d’outils, navigation en arborescence, groupes de fichiers, options progressives et barre d’état.
- 08:49 Migration de Zen vers WinUI 3 et Windows App SDK : interface Windows 11 avec Mica, navigation native, cartes système, retours d’état et exécutable x64 autonome testé au démarrage.
- 08:55 Correction de la publication autonome : intégration des ressources XAML requises et livraison de l’exécutable Windows 11 local dans le dossier Zen-Windows11, validé ouvert et actif.
- 09:02 Refonte de l’alignement et du parcours de conversion Windows 11 : colonne de travail unifiée, fichiers regroupés dans un seul flux, action principale visible immédiatement et désactivée tant qu’aucun fichier n’est choisi.
- 09:03 Clarification des actions principales : les conversions utilisent désormais « Convertir », tandis que les générations d’images et de PDF utilisent « Créer ».
- 09:10 Nouvelle répartition large inspirée des recommandations officielles Windows 11 : navigation adaptative, gouttières de 32 pixels, espace fichiers majoritaire et panneau latéral de résumé/action occupant toute la largeur utile.
- 09:24 Alignement exact des panneaux « Fichiers » et « Résumé » : suppression de l’espacement résiduel des options masquées afin d’obtenir les mêmes bords supérieur et inférieur.
- 18:57 Remplacement de la reconstruction PDF→DOCX en texte brut par Microsoft Word PDF Reflow sur un thread STA dédié, avec incorporation des polices et fermeture sûre des objets COM.
- 18:57 Ajout d’une validation automatique et ciblée de l’avertissement d’import PDF de Word, limitée à l’instance temporaire créée par Zen et sans modification persistante des préférences Office.
- 18:57 Conservation renforcée de la mise en page éditable : restauration des retours de ligne visuels du PDF sans perdre les styles de runs, et recréation des hyperliens e-mail et web dans le DOCX.
- 18:57 Validation complète avec test-2.pdf depuis l’interface native : DOCX valide, deux polices incorporées, dix-huit retours de ligne, lien mailto actif et aucun processus Word orphelin.
- 18:57 Refactorisation Clean Code du moteur : FileProcessor réduit à un routeur de 61 lignes et séparation des responsabilités en processeurs PDF, DOCX, OCR, images, restauration de mise en page et automatisation Word.

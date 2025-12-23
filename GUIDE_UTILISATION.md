# Guide d'utilisation - Excel Processor

## Description

Cette application permet de traiter deux fichiers Excel :
- **CLIENT.xlsx** : Fichier contenant les commandes clients
- **BackLog.xlsx** : État des commandes dans votre système

## Fonctionnalités

### 1. Confirmation des informations
L'application remplit automatiquement les colonnes T, U et V du fichier CLIENT :
- **Colonne T** : Numéro AR fournisseur (confirmé depuis BACKLOG)
- **Colonne U** : Nouveau délai confirmé par le fournisseur (depuis ConfirmedDate)
- **Colonne V** : Nouveau commentaire fournisseur (depuis Comment)

Les cellules mises à jour sont surlignées en vert clair.

### 2. Génération des tableaux de devancement
L'application identifie automatiquement les demandes de devancement :
- Une demande de devancement = ligne où la date souhaitée < date promise
- Génère un fichier Excel avec :
  - **Feuille récapitulative** : Liste de toutes les demandes
  - **Feuilles de détail** : Une feuille par demande (max 10)

## Installation sur Windows 11 (sans droits administrateur)

### Option 1 : Utiliser l'exécutable (RECOMMANDÉ)

1. **Télécharger** le fichier `ExcelProcessor.exe`
2. **Copier** le fichier sur votre PC Windows (n'importe quel dossier)
3. **Double-cliquer** sur `ExcelProcessor.exe`
4. **C'est tout !** Aucune installation nécessaire

### Option 2 : Utiliser Python portable (si l'exe ne fonctionne pas)

1. Télécharger WinPython (version sans installation) : https://winpython.github.io/
2. Extraire WinPython dans un dossier
3. Copier les fichiers :
   - `excel_processor.py`
   - `requirements.txt`
4. Ouvrir "WinPython Command Prompt"
5. Installer les dépendances :
   ```
   pip install -r requirements.txt
   ```
6. Lancer l'application :
   ```
   python excel_processor.py
   ```

## Utilisation de l'application

### Étape 1 : Lancer l'application
- Double-cliquez sur `ExcelProcessor.exe`
- Une fenêtre s'ouvre avec l'interface graphique

### Étape 2 : Sélectionner les fichiers
1. **Fichier CLIENT** : Cliquez sur "Parcourir..." et sélectionnez votre fichier CLIENT.xlsx
2. **Fichier BACKLOG** : Cliquez sur "Parcourir..." et sélectionnez votre fichier BackLog.xlsx
3. **Dossier de sortie** : (Optionnel) Choisissez où sauvegarder les résultats
   - Par défaut : même dossier que le fichier CLIENT

### Étape 3 : Choisir les traitements
- ☑ **Confirmer les informations** : Mettre à jour les colonnes T, U, V
- ☑ **Générer les tableaux de devancement** : Créer le fichier de devancement

Les deux options sont cochées par défaut.

### Étape 4 : Lancer le traitement
1. Cliquez sur **"Lancer le traitement"**
2. Attendez la fin du traitement (la barre de progression s'anime)
3. Les résultats s'affichent dans la zone de log

### Étape 5 : Récupérer les résultats
Deux fichiers sont créés dans le dossier de sortie :
- `CLIENT_confirme_YYYYMMDD_HHMMSS.xlsx` : Fichier CLIENT avec colonnes T, U, V complétées
- `Devancements_YYYYMMDD_HHMMSS.xlsx` : Fichier contenant les demandes de devancement

## Format des fichiers de sortie

### CLIENT_confirme_*.xlsx
- Identique au fichier CLIENT original
- Colonnes T, U, V complétées avec les données du BACKLOG
- Cellules mises à jour en vert clair

### Devancements_*.xlsx

#### Feuille "Récapitulatif"
Colonnes :
- Symbole
- Designation
- Numero AR
- Date promise
- Date souhaitée
- Jours de devancement (calculé automatiquement)
- Fournisseur
- Quantite

#### Feuilles "Dev_1_*", "Dev_2_*", etc.
Détails de chaque demande :
- Informations de la commande CLIENT
- Informations correspondantes du BACKLOG
- Calculs automatiques

## Correspondance entre fichiers

L'application utilise le champ **"Numero AR fournisseur"** (colonne T du CLIENT) pour faire la correspondance avec le champ **"OrderNo"** du BACKLOG.

## Dépannage

### L'exécutable ne se lance pas
- Vérifiez que vous avez les droits de lecture/écriture sur le dossier
- Essayez de copier l'exe dans un autre dossier (ex: Bureau)
- Utilisez l'Option 2 (Python portable)

### Message "Fichier corrompu" ou "Erreur Excel"
- Vérifiez que les fichiers CLIENT et BACKLOG sont bien au format .xlsx
- Essayez d'ouvrir et sauvegarder les fichiers dans Excel avant traitement

### Aucune demande de devancement trouvée
- Vérifiez que le fichier CLIENT contient bien :
  - Colonne "Date livraison souhaitee"
  - Colonne "Date initiale promise"
- Les dates doivent être au format date Excel (pas du texte)

### Aucune information confirmée
- Vérifiez que les numéros AR dans CLIENT correspondent aux OrderNo dans BACKLOG
- Format attendu : "0000150733" (avec les zéros en début)

## Support

Pour toute question ou problème :
1. Consultez la zone de log dans l'application pour les messages d'erreur
2. Vérifiez que vos fichiers respectent le format attendu
3. Essayez avec un sous-ensemble de données pour isoler le problème

## Informations techniques

### Fichiers générés
- Format : Excel (.xlsx)
- Encodage : UTF-8
- Compatible : Excel 2007 et versions ultérieures

### Limitations
- Fichiers BACKLOG : Peut traiter jusqu'à 50 000 lignes
- Fichiers CLIENT : Peut traiter jusqu'à 10 000 lignes
- Feuilles de détail : Maximum 10 demandes de devancement détaillées

### Sécurité
- Aucune donnée n'est envoyée sur Internet
- Traitement 100% local sur votre ordinateur
- Les fichiers originaux ne sont jamais modifiés

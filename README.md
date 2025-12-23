# Excel Processor - Traitement CLIENT / BACKLOG

Application Windows standalone pour traiter les fichiers CLIENT et BACKLOG sans nécessiter de droits administrateur.

## 🎯 Fonctionnalités

### 1. Confirmation des informations
- Remplit automatiquement les colonnes T, U, V du fichier CLIENT à partir du BACKLOG
- Surligne en vert les cellules mises à jour
- Génère un nouveau fichier `CLIENT_confirme_*.xlsx`

### 2. Tableaux de devancement
- Identifie automatiquement les demandes de devancement
- Crée un fichier Excel avec une feuille récapitulative
- Génère des feuilles de détail pour chaque demande (max 10)
- Calcule automatiquement le nombre de jours de devancement

## 📦 Contenu du package

```
excel-processor/
├── excel_processor.py      # Script Python principal
├── requirements.txt        # Dépendances Python
├── build_exe.bat          # Script de compilation Windows
├── build_exe.sh           # Script de compilation Linux/Mac
├── GUIDE_UTILISATION.md   # Guide complet en français
└── README.md              # Ce fichier
```

## 🚀 Installation et utilisation

### Option A : Utiliser l'exécutable pré-compilé (PLUS SIMPLE)

**Si vous avez reçu le fichier `ExcelProcessor.exe`** :
1. Copiez `ExcelProcessor.exe` sur votre PC Windows
2. Double-cliquez pour lancer
3. Aucune installation nécessaire !

### Option B : Compiler vous-même l'exécutable

**Prérequis** :
- Python 3.8 ou supérieur
- Connexion Internet (pour télécharger les dépendances)

**Sur Windows** :
```batch
1. Télécharger et installer Python : https://www.python.org/downloads/
   ⚠️ Cochez "Add Python to PATH" pendant l'installation

2. Double-cliquer sur build_exe.bat

3. L'exécutable sera créé dans dist\ExcelProcessor.exe
```

**Sur Linux/Mac** :
```bash
chmod +x build_exe.sh
./build_exe.sh
```

### Option C : Exécuter directement avec Python

```bash
# Installer les dépendances
pip install -r requirements.txt

# Lancer l'application
python excel_processor.py
```

## 💻 Utilisation

1. **Lancer l'application**
   - Double-cliquez sur `ExcelProcessor.exe`

2. **Sélectionner les fichiers**
   - Fichier CLIENT : Votre fichier CLIENT.xlsx
   - Fichier BACKLOG : Votre fichier BackLog.xlsx
   - Dossier de sortie : Où sauvegarder les résultats (optionnel)

3. **Choisir les traitements**
   - ☑ Confirmer les informations (colonnes T, U, V)
   - ☑ Générer les tableaux de devancement

4. **Lancer le traitement**
   - Cliquez sur "Lancer le traitement"
   - Attendez la fin (barre de progression)
   - Récupérez vos fichiers dans le dossier de sortie

## 📊 Fichiers générés

### `CLIENT_confirme_YYYYMMDD_HHMMSS.xlsx`
- Copie du fichier CLIENT avec colonnes T, U, V complétées
- Données extraites du BACKLOG via le numéro AR fournisseur
- Cellules mises à jour surlignées en vert

### `Devancements_YYYYMMDD_HHMMSS.xlsx`
- **Feuille "Récapitulatif"** : Liste toutes les demandes de devancement
- **Feuilles détails** : Une par demande avec informations complètes

## 🔧 Correspondance des données

L'application utilise :
- **CLIENT** : Colonne T "Numero AR fournisseur"
- **BACKLOG** : Colonne "OrderNo"

Ces deux champs doivent correspondre pour que la confirmation fonctionne.

## ⚙️ Configuration requise

### Pour exécuter l'application
- Windows 11 (ou Windows 10, 8, 7)
- **Aucun droit administrateur requis**
- 50 MB d'espace disque
- Fichiers Excel au format .xlsx

### Pour compiler l'exécutable
- Python 3.8+
- pip (gestionnaire de packages Python)
- Connexion Internet

## 📝 Format des fichiers

### Fichier CLIENT attendu
Colonnes requises :
- `Symbole`
- `Designation`
- `Numero AR fournisseur` (colonne T)
- `Nouveau delai confirme par le fournisseur` (colonne U)
- `Nouveau Commentaire fournisseur` (colonne V)
- `Date livraison souhaitee`
- `Date initiale promise`

### Fichier BACKLOG attendu
Colonnes requises :
- `OrderNo`
- `ConfirmedDate`
- `Comment`
- `OrderedQuantity`
- `RemainingQuantity`
- `DepartureDate`

## ❓ Dépannage

### L'exécutable ne se lance pas
- Vérifiez les droits de lecture/écriture sur le dossier
- Essayez de le copier sur le Bureau
- Désactivez temporairement l'antivirus (peut bloquer les exe non signés)

### Aucune information confirmée
- Vérifiez que les numéros AR dans CLIENT correspondent aux OrderNo dans BACKLOG
- Format : "0000150733" (avec zéros au début)
- Vérifiez qu'il n'y a pas d'espaces avant/après les numéros

### Aucune demande de devancement trouvée
- Vérifiez que les colonnes de dates existent
- Les dates doivent être au format date Excel (pas du texte)
- La date souhaitée doit être < date promise

## 📚 Documentation complète

Consultez `GUIDE_UTILISATION.md` pour :
- Instructions détaillées étape par étape
- Captures d'écran de l'interface
- Exemples de fichiers générés
- Résolution de problèmes avancée

## 🔒 Sécurité et confidentialité

- ✅ Traitement 100% local (aucune donnée envoyée sur Internet)
- ✅ Les fichiers originaux ne sont jamais modifiés
- ✅ Nouveaux fichiers créés avec timestamp unique
- ✅ Code source ouvert et vérifiable

## 📄 License

Ce logiciel est fourni "tel quel" sans garantie d'aucune sorte.
Utilisation libre pour un usage personnel ou professionnel.

## 🛠️ Technologies utilisées

- **Python 3** : Langage de programmation
- **tkinter** : Interface graphique (inclus dans Python)
- **pandas** : Manipulation de données Excel
- **openpyxl** : Lecture/écriture de fichiers Excel
- **PyInstaller** : Compilation en exécutable standalone

## 📞 Support

Pour toute question :
1. Consultez le `GUIDE_UTILISATION.md`
2. Vérifiez la zone de log dans l'application pour les messages d'erreur
3. Testez avec un petit échantillon de données

## 🔄 Versions

### v1.0 (2024)
- Première version
- Confirmation des informations (colonnes T, U, V)
- Génération des tableaux de devancement
- Interface graphique complète
- Compilation en exécutable Windows standalone

---

**Développé pour fonctionner sans droits administrateur sur Windows 11**

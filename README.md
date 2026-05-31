# Migration Gmail (MBOX) vers Outlook (PST avec catégories)

Ce script permet de convertir un fichier `.mbox` (export Google Takeout) ou un dossier de fichiers `.mbox` en un fichier `.pst` compatible avec Outlook Desktop, tout en transformant les étiquettes Gmail (`labels`) en **catégories Outlook** avec conservation de leurs couleurs d'origine d'un PC à un autre.

---

## 🚀 Fonctionnalités

### Conversion et Métadonnées
- **Conversion des labels en catégories** : préserve l'organisation Gmail sans dupliquer les messages.
- **Conservation des métadonnées** : Sujet, Expéditeur, Destinataire, Date, Pièces jointes.
- **Support HTML et UTF-8** : préserve la mise en forme et les caractères spéciaux.
- **Décodage MIME complet** : noms d'expéditeurs avec accents correctement affichés.
- **Threading des conversations** : regroupe automatiquement les messages d'un même fil dans Outlook (via les en-têtes `References` et `In-Reply-To`).

### Rendu et Performance (Optimisé)
- **Liaison anticipée (Early Binding COM)** : utilise `EnsureDispatch` pour accélérer de façon significative les interactions avec Outlook COM (avec fallback transparent vers `Dispatch` standard).
- **Rendu parfait des images inline** : applique les propriétés MAPI complexes (`PR_ATTACH_FLAGS`, `PR_ATTACH_MIME_TAG`, `PR_ATTACHMENT_HIDDEN`) pour que les images s'affichent correctement dans le corps HTML sans apparaître en doublon dans la barre des pièces jointes d'Outlook.
- **Parser MBOX standard** : lecture robuste via la bibliothèque standard de Python pour éviter les troncatures d'images.

### Flexibilité et Volume
- **Mode Batch (Multi-MBOX)** : accepte un répertoire contenant plusieurs fichiers `.mbox` (utile pour les exports Google Takeout volumineux splittés) et les importe à la suite dans le même fichier PST.
- **Interface Graphique Moderne (GUI)** : interface Tkinter/Ttk complète et élégante pour configurer et lancer la migration en quelques clics sans passer par le terminal.
- **Limitation optionnelle** : paramètre `--limit` pour tester la migration sur un nombre réduit de messages.

### Robustesse et Reprise
- **Déduplication robuste par Message-ID** : filtre automatiquement les doublons. La liste des e-mails déjà migrés est sauvegardée dans l'état de reprise pour éviter les doublons même après une interruption.
- **Reprise sur interruption** : sauvegarde automatique de l'état tous les 100 messages.
- **Arrêt gracieux (Ctrl+C ou bouton Pause)** : sauvegarde immédiate de l'état avant fermeture.
- **Rapport des erreurs** : fichier `problem_messages.json` listant les messages problématiques.

---

## 🛠️ Prérequis

1. **Windows** avec **Microsoft Outlook (Classic)** installé.
2. Outlook doit être configuré avec un profil actif.
   * *Astuce : Pour utiliser Outlook sans lui associer d'adresse e-mail (mode local seul), ouvrez l'invite de commande (Win+R) et lancez :*
     ```text
     outlook.exe /PIM "MonProfilHorsLigne"
     ```
3. **Python 3.x** installé.
4. Bibliothèques Python requises :
   ```bash
   pip install pywin32 tqdm
   ```

---

## 📖 Utilisation

### 1. Mode Interface Graphique (GUI) — Recommandé
Double-cliquez sur le script `mbox_to_pst.py` ou lancez-le simplement dans votre terminal sans argument :
```bash
python mbox_to_pst.py
```
*L'interface graphique vous permettra de sélectionner vos fichiers, configurer vos options et suivre la progression en temps réel avec un journal de logs intégré.*

### 2. Mode Ligne de commande (CLI)
#### Commande de base (Fichier unique) :
```bash
python mbox_to_pst.py "chemin/vers/fichier.mbox" "chemin/vers/sortie.pst"
```

#### Traitement par lots (Batch dossier) :
```bash
python mbox_to_pst.py "chemin/vers/dossier_mboxes" "chemin/vers/sortie.pst"
```

#### Avec limitation de messages (pour tests) :
```bash
python mbox_to_pst.py "fichier.mbox" "sortie.pst" --limit 100
```

### Options CLI disponibles

| Option | Description |
|--------|-------------|
| `--folder "Nom"` | Nom du dossier racine dans le PST (défaut: "Gmail Archive") |
| `--limit N` | Limite le traitement à N messages par fichier |
| `--no-resume` | Ignore l'état précédent et recommence depuis le début |
| `--gui` | Force l'ouverture de l'interface graphique |

---

## 📦 Fichiers générés

| Fichier | Description |
|---------|-------------|
| `migration.log` | Journal détaillé des opérations de migration |
| `migration_state.json` | État pour la reprise après interruption (contient l'index et les Message-IDs déjà vus) |
| `problem_messages.json` | Liste des messages avec erreurs (pièces jointes corrompues, etc.) |
| `categories_config.json` | Configuration exportée des couleurs de catégories (généré par le script d'export) |

---

## 🏷️ Gestion et Portabilité des Catégories (PowerShell)

Les catégories Outlook et leurs couleurs sont normalement liées au profil local de Windows. Ces scripts PowerShell (encodés en UTF-8 avec BOM) permettent de transférer vos catégories et couleurs de façon transparente.

### `export_categories.ps1` — Exportation (PC Source)
Exporte la liste globale des catégories de votre profil Outlook (Nom, Couleur indexée de 0 à 25, Raccourcis) dans un fichier `categories_config.json`.

### `sync_categories.ps1` — Synchronisation et Restauration (PC Cible)
Scanne récursivement le fichier PST importé pour récupérer les catégories appliquées aux messages et les recréer dans le profil du nouveau PC. **Si le fichier `categories_config.json` est présent dans le même répertoire, il restaure automatiquement les couleurs et raccourcis d'origine.**

### `manage_categories.ps1` — Nettoyage complet
Permet de lister, supprimer sélectivement ou supprimer en masse les catégories de votre profil Outlook (du catalogue seul, ou du catalogue et de tous les messages pour une remise à zéro).

---

## 📋 Procédure Recommandée de Bout en Bout

### Étape 1 : Migration et configuration (PC Source - PC 1)
1. Exécutez la migration (GUI ou CLI). Les étiquettes Gmail sont importées dans le fichier PST en tant que catégories Outlook.
2. Ouvrez Outlook sur le PC 1 et attribuez les couleurs de votre choix à vos nouvelles catégories (Ruban Outlook > **Classer** > **Toutes les catégories**).
3. Ouvrez PowerShell dans le dossier du projet et exécutez le script d'exportation :
   ```powershell
   PowerShell.exe -ExecutionPolicy Bypass -File .\export_categories.ps1
   ```
   *(Cela génère le fichier `categories_config.json`)*.

### Étape 2 : Transfert (PC 1 ➔ PC 2)
Copiez sur votre support de transfert vers le PC cible :
- Le fichier **`.pst`** migré.
- Le fichier **`categories_config.json`**.
- Les scripts **`sync_categories.ps1`** et **`manage_categories.ps1`**.

### Étape 3 : Importation et Synchronisation (PC Cible - PC 2)
1. Dans Outlook sur le PC 2, ouvrez le PST : **Fichier** > **Ouvrir et exporter** > **Ouvrir le fichier de données Outlook**. *(Les étiquettes de catégories apparaissent en gris/blanc)*.
2. *(Optionnel)* Si d'anciens essais ont créé des catégories fantômes, lancez `manage_categories.ps1` avec l'option **`[A]`** (Tout supprimer) puis le mode **`[2]`** (Nettoyage complet) pour nettoyer le profil.
3. Lancez la restauration des couleurs et catégories en exécutant :
   ```powershell
   PowerShell.exe -ExecutionPolicy Bypass -File .\sync_categories.ps1
   ```
   *(Sélectionnez le numéro de votre PST importé. Les catégories seront créées dans Outlook avec leurs couleurs d'origine et ré-appliquées proprement).*

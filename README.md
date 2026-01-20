# Migration Gmail (MBOX) vers Outlook (PST avec catégories) 

Ce script permet de convertir un fichier `.mbox` (export Google Takeout) en un fichier `.pst` compatible avec Outlook Desktop, tout en transformant les étiquettes Gmail (`labels`) en **catégories Outlook**.

## 🚀 Fonctionnalités

### Conversion et Métadonnées
- ✅ **Conversion des labels en catégories** : préserve l'organisation Gmail sans dupliquer les messages
- ✅ **Conservation des métadonnées** : Sujet, Expéditeur, Destinataire, Date, Pièces jointes
- ✅ **Support du HTML et de l'UTF-8** : préserve la mise en forme et les caractères spéciaux
- ✅ **Décodage MIME complet** : noms d'expéditeurs avec accents correctement affichés

### Performance et Fichiers Volumineux
- ✅ **Parser MBOX streaming** : lecture par blocs de 1 Mo au lieu du chargement mémoire complet
- ✅ **Optimisé pour les gros volumes** : testé avec des fichiers jusqu'à 10 Go
- ✅ **Barre de progression en temps réel** : affichage fluide basé sur la position dans le fichier

### Gestion des Doublons
- ✅ **Déduplication par Message-ID** : évite l'import de messages en double (fréquent avec les exports Gmail multi-labels)
- ✅ **Compteur de doublons** : affiche le nombre de messages ignorés à la fin

### Robustesse et Reprise
- ✅ **Reprise sur interruption** : sauvegarde automatique de l'état tous les 100 messages
- ✅ **Arrêt gracieux (Ctrl+C)** : sauvegarde immédiate de l'état avant fermeture
- ✅ **Rapport des erreurs** : fichier `problem_messages.json` listant les messages problématiques

### Qualité des Messages Importés
- ✅ **Corrections du statut Brouillon** : les messages n'apparaissent plus comme brouillons dans Outlook
- ✅ **Dates d'envoi préservées** : affichage correct des dates originales
- ✅ **Mise à jour automatique des catégories Outlook** : coloration immédiate disponible

## 🛠️ Prérequis

1.  **Windows** avec **Microsoft Outlook** installé
2.  Démarer Outlook sans comptes avec la commande
    --> Alt+R "outlook.exe /PIM nom_de_profile" (crée le profile)
    --> Alt+R "outlook.exe /profile nom_de_profile" (ouvre le profile)
2.  **Python 3.x**
3.  Bibliothèques Python :
    ```bash
    pip install pywin32 tqdm
    ```

## 📖 Utilisation

### Commande de base
```bash
python mbox_to_pst.py "chemin/vers/fichier.mbox" "chemin/vers/sortie.pst"
```

### Avec limitation de messages (pour tests)
```bash
python mbox_to_pst.py "fichier.mbox" "sortie.pst" --limit 100
```

### Options disponibles

| Option | Description |
|--------|-------------|
| `--folder "Nom"` | Nom du dossier racine dans le PST (défaut: "Gmail Archive") |
| `--limit N` | Limite le traitement à N messages (utile pour les tests) |
| `--no-resume` | Ignore l'état précédent et recommence depuis le début |

## 🛑 Arrêter et Reprendre

- **Arrêter proprement** : Appuyez sur `Ctrl+C` → l'état est sauvegardé immédiatement
- **Reprendre** : Relancez la même commande → reprise automatique au dernier message traité
- **Recommencer à zéro** : Supprimez `migration_state.json`

## 📦 Fichiers générés

| Fichier | Description |
|---------|-------------|
| `migration.log` | Journal détaillé des opérations |
| `migration_state.json` | État pour la reprise après interruption |
| `problem_messages.json` | Liste des messages avec erreurs (pièces jointes trop volumineuses, etc.) |

## ⚠️ Notes importantes

- **Outlook doit être installé** : le script utilise l'interface COM native
- **Vitesse** : ~2-5 messages/seconde (les fichiers de 10 Go peuvent prendre plusieurs heures)
- **Ne pas fermer Outlook** pendant l'exécution du script
- **Doublons Gmail** : automatiquement filtrés grâce à la déduplication par Message-ID

---

## 🏷️ Scripts de Gestion des Catégories Outlook (PowerShell)

Après migration du PST vers un autre PC, les catégories peuvent ne pas être reconnues par Outlook (affichage blanc/gris). Ces scripts permettent de synchroniser et gérer les catégories.

### `sync_categories.ps1` — Synchronisation et Réparation

**Utilité :** Importer les catégories d'un PST vers la "Master Category List" d'Outlook et réparer les étiquettes fusionnées.

**Fonctionnement :**
1. Parcourt récursivement tous les messages du PST sélectionné.
2. Découpe les étiquettes fusionnées (ex: `"A; B; C"` → trois catégories).
3. Ajoute chaque catégorie à la Master List si elle n'existe pas.
4. Ré-enregistre les catégories proprement sur chaque message.

**Cas d'usage :**
- Post-migration sur un nouveau PC
- Réparation des couleurs de catégories

---

### `manage_categories.ps1` — Gestion et Nettoyage

**Utilité :** Lister, supprimer sélectivement ou en masse les catégories Outlook.

**Options :**
- **[N]** : Supprimer par numéro(s) (ex: `1,5,10,20`)
- **[A]** : Supprimer TOUTES les catégories
- **[Q]** : Quitter

**Modes de suppression :**
- **[1]** : Catalogue Outlook uniquement
- **[2]** : Catalogue + tous les messages (Nettoyage Complet)

**Cas d'usage :**
- Nettoyer les catégories fusionnées créées par erreur
- Remise à zéro avant resynchronisation
- Supprimer des catégories obsolètes (Spam, Forums, etc.)

---

### 📋 Procédure Recommandée (PC de destination)

```powershell
# 1. Nettoyer (si nécessaire)
PowerShell.exe -ExecutionPolicy Bypass -File .\manage_categories.ps1
# → [A] pour tout sélectionner → [2] nettoyage complet → [0] tous les comptes

# 2. Synchroniser
PowerShell.exe -ExecutionPolicy Bypass -File .\sync_categories.ps1
# → Sélectionner le PST → Laisser réparer

# 3. Dans Outlook : Classer > Toutes les catégories > Attribuer les couleurs
```

> **Note :** Les scripts sont encodés en UTF-8 avec BOM pour gérer les caractères accentués.

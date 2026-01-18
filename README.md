# Migration Gmail (MBOX) vers Outlook (PST avec catégories)

Ce script permet de convertir un fichier `.mbox` (export Google Takeout) en un fichier `.pst` compatible avec Outlook Desktop, tout en transformant les étiquettes Gmail (`labels`) en **catégories Outlook**.

## 🚀 Fonctionnalités

- ✅ **Conversion des labels en catégories** : préserve l'organisation Gmail sans dupliquer les messages.
- ✅ **Gestion des gros volumes** : optimisé pour des fichiers jusqu'à 10 Go (traitement itératif).
- ✅ **Conservation des métadonnées** : Sujet, Expéditeur, Destinataire, Date, Pièces jointes.
- ✅ **Support du HTML et de l'UTF-8** : préserve la mise en forme et les caractères spéciaux.
- ✅ **Reprise sur interruption** : En cas de plantage ou d'arrêt manuel, le script peut reprendre là où il s'est arrêté.
- ✅ **Mise à jour de la liste des catégories** : Ajoute automatiquement les nouveaux labels à la liste des catégories Outlook pour une coloration immédiate.

## 🛠️ Prérequis

1.  **Windows** avec **Microsoft Outlook** installé.
2.  **Python 3.x** installé.
3.  Bibliothèque `pywin32` installée :
    ```bash
    pip install pywin32
    ```

## 📖 Utilisation

1.  Ouvrez un terminal (PowerShell ou Command Prompt).
2.  Lancez le script avec le chemin de votre fichier MBOX et le chemin du fichier PST souhaité :

```bash
python mbox_to_pst.py "E:\Sauveguarde_Messages_GMAIL\Tous les messages, y compris ceux du dossier Spam -002.mbox" "E:\Sauveguarde_Messages_GMAIL\Takeout\Mail\archive_outlook.pst" --limit 50
```

### Options supplémentaires :

- `--folder "Archive Gmail"` : Permet de spécifier le nom du dossier racine dans le PST (par défaut: "Gmail Archive").
- `--no-resume` : Force le script à recommencer depuis le début (ignore l'état précédent).

## ⚠️ Notes importantes

- **Outlook doit être installé** sur la machine car le script utilise l'interface COM d'Outlook pour créer le fichier PST de manière native et fiable.
- **Vitesse** : L'interface COM d'Outlook peut être lente (~2-5 messages par seconde). Pour un fichier de 10 Go (potentiellement 100 000+ emails), le traitement peut durer plusieurs heures.
- **Stabilité** : Ne fermez pas Outlook pendant l'exécution du script. Le script créera une instance d'Outlook en arrière-plan si nécessaire.

## 📦 Fichiers générés

- `mbox_to_pst.py` : Le script principal.
- `migration.log` : Journal détaillé des opérations et des erreurs éventuelles.
- `migration_state.json` : Fichier temporaire permettant la reprise après interruption.

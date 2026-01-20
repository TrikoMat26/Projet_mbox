import codecs
import os

files = [
    r'e:\Programmation\Projet_mbox_pst\sync_categories.ps1',
    r'e:\Programmation\Projet_mbox_pst\manage_categories.ps1'
]

for path in files:
    if os.path.exists(path):
        print(f"Traitement de : {path}")
        # Lire en UTF-8 (sans BOM ou avec)
        with open(path, 'rb') as f:
            raw = f.read()
        
        # Détection basique pour éviter les erreurs de décodage
        content = raw.decode('utf-8-sig') # Gère avec ou sans BOM existant
        
        # Réenregistrer avec UTF-8-SIG (BOM)
        with codecs.open(path, 'w', 'utf-8-sig') as f:
            f.write(content)
        print(f"  OK : Enregistré en UTF-8 avec BOM.")
    else:
        print(f"  Erreur : Fichier non trouvé : {path}")

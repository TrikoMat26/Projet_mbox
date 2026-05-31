import codecs
import os

# Resolve paths relative to this script's directory for portability
script_dir = os.path.dirname(os.path.abspath(__file__))

files = [
    os.path.join(script_dir, 'sync_categories.ps1'),
    os.path.join(script_dir, 'manage_categories.ps1'),
    os.path.join(script_dir, 'export_categories.ps1')
]

for path in files:
    if os.path.exists(path):
        print(f"Processing: {path}")
        # Read file contents in UTF-8 (handling existing BOM dynamically)
        with open(path, 'rb') as f:
            raw = f.read()
        
        # Decode using utf-8-sig to strip BOM if present
        content = raw.decode('utf-8-sig')
        
        # Rewrite file with UTF-8-SIG (forces Windows PowerShell to recognize accents)
        with codecs.open(path, 'w', 'utf-8-sig') as f:
            f.write(content)
        print(f"  Success: Saved with UTF-8 BOM encoding.")
    else:
        print(f"  Error: File not found at {path}")

# Script de synchronisation des catégories Outlook
# Ce script scanne un fichier PST et ajoute toutes les catégories trouvées à votre liste principale Outlook.

$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

# 1. Demander le nom du PST à scanner
Write-Host "--- Synchronisation des Catégories Outlook ---" -ForegroundColor Cyan
Write-Host "Voici vos dossiers Outlook actuels :"
$folders = $namespace.Folders
for ($i = 1; $i -le $folders.Count; $i++) {
    Write-Host "[$i] $($folders.Item($i).Name)"
}

$choice = Read-Host "Entrez le numéro du PST (ex: Gmail Archive) à scanner"
if ($choice -le 0 -or $choice -gt $folders.Count) {
    Write-Host "Choix invalide." -ForegroundColor Red
    exit
}

$targetPST = $folders.Item([int]$choice)
Write-Host "Scan du dossier : $($targetPST.Name)..." -ForegroundColor Yellow

# 2. Récupérer la liste des catégories existantes
$masterCategories = @{}
foreach ($cat in $namespace.Categories) {
    $masterCategories[$cat.Name] = $true
}

# 3. Fonction récursive pour scanner les dossiers
function Scan-FolderCategories($folder) {
    Write-Host "  Lecture de : $($folder.Name) ($($folder.Items.Count) messages)" -ForegroundColor Gray
    
    foreach ($item in $folder.Items) {
        if ($item.Categories) {
            # Outlook can use comma or semicolon depending on locale
            $cats = $item.Categories -split '[;,]'
            foreach ($c in $cats) {
                $cleanCat = $c.Trim()
                if ($cleanCat -and -not $masterCategories.ContainsKey($cleanCat)) {
                    try {
                        Write-Host "    [NEW] Ajout de la catégorie : $cleanCat" -ForegroundColor Green
                        $namespace.Categories.Add($cleanCat)
                        $masterCategories[$cleanCat] = $true
                    }
                    catch {
                        Write-Host "    Erreur lors de l'ajout de $cleanCat" -ForegroundColor Red
                    }
                }
            }
        }
    }
    
    foreach ($subFolder in $folder.Folders) {
        Scan-FolderCategories $subFolder
    }
}

# Lancer le scan
Scan-FolderCategories $targetPST

Write-Host "`nTerminé !" -ForegroundColor Cyan
Write-Host "Vous pouvez maintenant attribuer des couleurs aux nouvelles catégories dans Outlook (Ruban Accueil -> Classer -> Toutes les catégories)."
$outlook.Quit()

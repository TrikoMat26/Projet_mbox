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

# 3. Fonction récursive pour scanner et REPARER les catégories
function Scan-FolderCategories($folder) {
    Write-Host "`nScan de : $($folder.Name) ($($folder.Items.Count) messages)" -ForegroundColor Yellow
    
    $count = 0
    foreach ($item in $folder.Items) {
        if ($item.Categories) {
            # 1. On découpe et on nettoie
            $rawCats = $item.Categories -split '[;,]'
            $cleanParts = New-Object System.Collections.Generic.List[string]
            
            foreach ($c in $rawCats) {
                $trimmed = $c.Trim()
                if ($trimmed) {
                    $cleanParts.Add($trimmed)
                    # 2. On s'assure que ça existe dans la Master List
                    if (-not $masterCategories.ContainsKey($trimmed)) {
                        try {
                            Write-Host "    [NEW] Ajout Master List : $trimmed" -ForegroundColor Green
                            $namespace.Categories.Add($trimmed)
                            $masterCategories[$trimmed] = $true
                        }
                        catch {}
                    }
                }
            }

            # 3. REPARATION : On ré-écrit les catégories proprement sur le message
            $newCatString = [string]::Join("; ", $cleanParts)
            if ($item.Categories -ne $newCatString) {
                try {
                    $item.Categories = $newCatString
                    $item.Save()
                    $count++
                    if ($count % 50 -eq 0) { Write-Host "  > $count messages réparés..." -ForegroundColor Gray }
                }
                catch {
                    Write-Host "  ! Erreur sur message : $($item.Subject)" -ForegroundColor Red
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

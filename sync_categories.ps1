# Script de synchronisation des catégories Outlook
# Ce script scanne un fichier PST, ajoute toutes les catégories trouvées à votre liste principale Outlook,
# et restaure leurs couleurs d'origine si le fichier categories_config.json est présent.

$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

# 1. Demander le nom du PST à scanner
Write-Host "--- Synchronisation et Restauration des Catégories Outlook ---" -ForegroundColor Cyan
Write-Host "Voici vos dossiers Outlook actuels :"
$folders = $namespace.Folders
for ($i = 1; $i -le $folders.Count; $i++) {
    Write-Host "[$i] $($folders.Item($i).Name)"
}

$choice = Read-Host "Entrez le numéro du PST à scanner"
if ($choice -le 0 -or $choice -gt $folders.Count) {
    Write-Host "Choix invalide." -ForegroundColor Red
    exit
}

$targetPST = $folders.Item([int]$choice)
Write-Host "Scan du dossier : $($targetPST.Name)..." -ForegroundColor Yellow

# 2. Charger les configurations de couleurs si présentes
$configPath = Join-Path $PSScriptRoot "categories_config.json"
$categoriesConfig = @{}
if (Test-Path $configPath) {
    Write-Host "`nFichier categories_config.json détecté. Chargement des couleurs et raccourcis..." -ForegroundColor Green
    try {
        $json = Get-Content -Raw -Path $configPath | ConvertFrom-Json
        foreach ($item in $json) {
            $categoriesConfig[$item.Name] = @{
                Color       = $item.Color
                ShortcutKey = $item.ShortcutKey
            }
        }
        Write-Host "  > $(($categoriesConfig.Keys).Count) configurations de couleurs chargées." -ForegroundColor Gray
    }
    catch {
        Write-Host "  ! Erreur de chargement du fichier JSON. Les couleurs par défaut seront utilisées." -ForegroundColor Yellow
    }
} else {
    Write-Host "`nAucun fichier categories_config.json trouvé dans ce répertoire." -ForegroundColor Yellow
    Write-Host "Les catégories seront recréées avec leurs couleurs par défaut." -ForegroundColor Gray
}

# 3. Récupérer la liste des catégories existantes sur ce PC
$masterCategories = @{}
foreach ($cat in $namespace.Categories) {
    $masterCategories[$cat.Name] = $true
}

# 4. Fonction récursive pour scanner et REPARER les catégories
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
                            
                            # Si on a la couleur dans la config, on l'applique
                            if ($categoriesConfig.ContainsKey($trimmed)) {
                                $cConfig = $categoriesConfig[$trimmed]
                                $colorVal = [int]$cConfig.Color
                                $shortcutVal = [int]$cConfig.ShortcutKey
                                
                                # Ajout avec couleur et raccourci d'origine
                                $namespace.Categories.Add($trimmed, $colorVal, $shortcutVal)
                                Write-Host "          -> Couleur rétablie (index $colorVal)" -ForegroundColor Green
                            }
                            else {
                                $namespace.Categories.Add($trimmed)
                            }
                            $masterCategories[$trimmed] = $true
                        }
                        catch {
                            # Fallback si l'ajout avec couleur échoue
                            try {
                                $namespace.Categories.Add($trimmed)
                                $masterCategories[$trimmed] = $true
                            } catch {}
                        }
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
Write-Host "Les catégories et couleurs ont été restaurées dans votre Outlook."
$outlook.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null

Write-Host "`nAppuyez sur Entrée pour quitter..."
Read-Host

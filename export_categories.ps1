# Script d'exportation des catégories Outlook avec leurs couleurs et raccourcis
# Ce script lit votre Master Category List Outlook locale et l'enregistre dans un fichier JSON.

Write-Host "--- Exportation des Catégories Outlook ---" -ForegroundColor Cyan

$outlook = $null
try {
    Write-Host "Initialisation d'Outlook..." -ForegroundColor Gray
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    
    $categories = $namespace.Categories
    Write-Host "Nombre de catégories trouvées : $($categories.Count)" -ForegroundColor Yellow
    
    if ($categories.Count -eq 0) {
        Write-Host "Aucune catégorie trouvée dans votre profil local. Export annulé." -ForegroundColor Red
        exit
    }
    
    $categoriesList = @()
    
    for ($i = 1; $i -le $categories.Count; $i++) {
        $cat = $categories.Item($i)
        Write-Host "  > Export : $($cat.Name) (Couleur index: $($cat.Color))" -ForegroundColor Gray
        
        $categoriesList += [PSCustomObject]@{
            Name        = $cat.Name
            Color       = [int]$cat.Color
            ShortcutKey = [int]$cat.ShortcutKey
        }
    }
    
    $outputPath = Join-Path $PSScriptRoot "categories_config.json"
    
    # Conversion en JSON et écriture
    $categoriesList | ConvertTo-Json -Depth 5 | Out-File -FilePath $outputPath -Encoding utf8
    
    Write-Host "`nExport réussi !" -ForegroundColor Green
    Write-Host "Fichier généré : categories_config.json" -ForegroundColor Cyan
    Write-Host "Veuillez copier ce fichier avec le script 'sync_categories.ps1' sur le PC de destination." -ForegroundColor Yellow
}
catch {
    Write-Host "`nERREUR LORS DE L'EXPORT : $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($outlook) {
        $outlook.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
    }
    Write-Host "`nAppuyez sur Entrée pour quitter..."
    Read-Host
}

# Script de gestion des catégories Outlook
# Ce script permet de lister et de supprimer les catégories de votre liste principale.

$outlook = $null
try {
    Write-Host "Initialisation d'Outlook..." -ForegroundColor Gray
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")

    while ($true) {
        $categories = $namespace.Categories
        Write-Host "`n--- Liste des Catégories Outlook ($($categories.Count)) ---" -ForegroundColor Cyan
        
        if ($categories.Count -eq 0) {
            Write-Host "Aucune catégorie trouvée dans votre profil Outlook."
        }
        else {
            for ($i = 1; $i -le $categories.Count; $i++) {
                $cat = $categories.Item($i)
                Write-Host "[$i] $($cat.Name)"
            }
        }

        Write-Host "`nOptions :" -ForegroundColor Yellow
        Write-Host "[N] Supprimer par numéro(s) (ex: 1,3,5)"
        Write-Host "[A] Supprimer TOUTES les catégories"
        Write-Host "[Q] Quitter"

        $action = Read-Host "Choisissez une action"

        if ($action -match "Q") {
            break
        }
        
        $selectedNames = @()
        
        if ($action -match "A") {
            if ($categories.Count -eq 0) {
                Write-Host "Rien à supprimer." -ForegroundColor Red
                continue
            }
            Write-Host "`nATTENTION : Vous allez supprimer TOUTES ($($categories.Count)) les catégories !" -ForegroundColor Red
            $confAll = Read-Host "Êtes-vous sûr ? (Tapez 'OUI' pour confirmer)"
            if ($confAll -ne "OUI") { Write-Host "Annulé."; continue }
            
            for ($i = 1; $i -le $categories.Count; $i++) {
                $selectedNames += $categories.Item($i).Name
            }
        }
        elseif ($action -match "N") {
            $selection = Read-Host "Entrez les numéros à supprimer (séparés par des virgules)"
            if (-not $selection) { continue }
            
            $indices = $selection.Split(",") | ForEach-Object { 
                $v = 0
                if ([int]::TryParse($_.Trim(), [ref]$v)) { [int]$v }
            } | Where-Object { $_ -ne $null } | Sort-Object -Descending -Unique

            if ($indices.Count -eq 0) { continue }

            Write-Host "`nVOUS AVEZ SÉLECTIONNÉ :" -ForegroundColor Yellow
            foreach ($i in $indices) {
                if ($i -ge 1 -and $i -le $categories.Count) {
                    $name = $categories.Item($i).Name
                    Write-Host "[$i] $name" -ForegroundColor Magenta
                    $selectedNames += $name
                }
            }
        }
        else {
            continue
        }

        if ($selectedNames.Count -eq 0) { continue }

        Write-Host "`nOptions de suppression :" -ForegroundColor Yellow
        Write-Host "[1] Supprimer du catalogue Outlook UNIQUEMENT"
        Write-Host "[2] Supprimer du catalogue ET des messages (Nettoyage Complet)"
        $mode = Read-Host "Votre choix (1 ou 2)"

        if ($mode -eq "2") {
            Write-Host "`nSélectionnez la cible du nettoyage :" -ForegroundColor Yellow
            Write-Host "[0] TOUS les comptes et fichiers PST (Recommandé)"
            $folders = $namespace.Folders
            for ($i = 1; $i -le $folders.Count; $i++) { Write-Host "[$i] $($folders.Item($i).Name)" }
            
            $fChoiceSelection = Read-Host "Choix (0 pour tout, ou numéro)"
            
            function Invoke-DeepClean($folder, $namesToRemove) {
                Write-Host "  Fouille : $($folder.Name)..." -ForegroundColor Gray
                $c = 0
                try {
                    foreach ($item in $folder.Items) {
                        if ($item.Categories) {
                            $initial = $item.Categories
                            $currentCats = $initial -split '[;,]' | ForEach-Object { $_.Trim() }
                            $filteredCats = @($currentCats | Where-Object { 
                                    $thisCat = $_
                                    -not ($namesToRemove | Where-Object { $_ -ieq $thisCat })
                                })
                            
                            $finalString = if ($filteredCats.Count -gt 0) { [string]::Join("; ", $filteredCats) } else { "" }
                            
                            if ($initial -ne $finalString) {
                                try {
                                    $item.Categories = $finalString
                                    $item.Save()
                                    $c++
                                }
                                catch {}
                            }
                        }
                    }
                }
                catch {}
                if ($c -gt 0) { Write-Host "    -> $c messages nettoyés dans $($folder.Name)." -ForegroundColor Green }
                foreach ($sub in $folder.Folders) { Invoke-DeepClean $sub $namesToRemove }
            }

            if ($fChoiceSelection -eq "0") {
                foreach ($store in $folders) { Invoke-DeepClean $store $selectedNames }
            }
            elseif ($fChoiceSelection -match '^\d+$' -and [int]$fChoiceSelection -ge 1 -and [int]$fChoiceSelection -le $folders.Count) {
                Invoke-DeepClean $folders.Item([int]$fChoiceSelection) $selectedNames
            }
        }

        # Suppression du Catalogue (Master List) - BOUCLE JUSQU'À SUCCÈS
        Write-Host "`nSuppression finale du catalogue Outlook..." -ForegroundColor Yellow
        $maxAttempts = 5
        $attempt = 0
        $remaining = $selectedNames.Clone()
        
        while ($remaining.Count -gt 0 -and $attempt -lt $maxAttempts) {
            $attempt++
            $stillRemaining = @()
            
            # Recharger la collection à chaque tentative
            $freshCategories = $namespace.Categories
            
            foreach ($n in $remaining) {
                $found = $false
                for ($i = 1; $i -le $freshCategories.Count; $i++) {
                    if ($freshCategories.Item($i).Name -ieq $n) {
                        $found = $true
                        break
                    }
                }
                
                if ($found) {
                    try {
                        $freshCategories.Remove($n)
                        Write-Host "  Retiré : $n" -ForegroundColor Green
                    }
                    catch {
                        $stillRemaining += $n
                    }
                }
            }
            
            # Vérifier ce qui reste vraiment
            Start-Sleep -Milliseconds 200
            $freshCategories = $namespace.Categories
            $remaining = @()
            foreach ($n in $selectedNames) {
                for ($i = 1; $i -le $freshCategories.Count; $i++) {
                    if ($freshCategories.Item($i).Name -ieq $n) {
                        $remaining += $n
                        break
                    }
                }
            }
            
            if ($remaining.Count -gt 0 -and $attempt -lt $maxAttempts) {
                Write-Host "  (Tentative $attempt : $($remaining.Count) catégorie(s) résistante(s), nouvelle tentative...)" -ForegroundColor Gray
            }
        }
        
        if ($remaining.Count -gt 0) {
            Write-Host "  Impossible de supprimer après $maxAttempts tentatives : $($remaining -join ', ')" -ForegroundColor Red
        }
        else {
            Write-Host "Opération terminée avec succès." -ForegroundColor Cyan
        }
    }
}
catch {
    Write-Host "`nERREUR CRITIQUE : $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Vérifiez qu'Outlook est bien ouvert sur ce PC." -ForegroundColor Yellow
}
finally {
    Write-Host "`nPress Entrée pour fermer cette fenêtre..."
    Read-Host
}

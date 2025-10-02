# PPT to PPTX Batch convert script
# Copié et modifié de la version XLSX
# Edité by Logan

$folderpath = "your_path"   # Dossier où le script agit
$filetype   = "*.ppt"                                           # Extensions de fichiers à cibler

# Chargement PowerPoint Interop, A commenté pour desactivé suivant la version d'excel installé
Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint

# Constante pour format PPTX
$ppSaveAsOpenXMLPresentation = [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsOpenXMLPresentation

# Lance PowerPoint
$ppt = New-Object -ComObject PowerPoint.Application

# Remise à zéro des compteurs de poids de fichiers, de nombre de fichiers convertis et d'erreurs
$taillePPT  = 0
$taillePPTX = 0
$nbConvertis = 0
$nbFails = 0

# Conversion des .ppt en .pptx
Get-ChildItem -Path $folderpath -Include $filetype -Recurse | ForEach-Object {
    $path = ($_.FullName).Substring(0, ($_.FullName).LastIndexOf("."))

    Write-Host "Conversion en cours : $($_.FullName)"

    try {
        # Récupération des métadonnées avant conversion
        $originalFile = $_
        $creationTime = $originalFile.CreationTime
        $lastWriteTime = $originalFile.LastWriteTime
        $lastAccessTime = $originalFile.LastAccessTime
        $owner = (Get-Acl $originalFile.FullName).Owner

        # Taille des fichiers avant conversion
        $taillePPT += $_.Length

        # Ouvrir et convertir
        $presentation = $ppt.Presentations.Open($_.FullName, $false, $false, $false)
        $newPath = $path + ".pptx"
        $presentation.SaveAs($newPath, $ppSaveAsOpenXMLPresentation)
        $presentation.Close()

        if (Test-Path $newPath) {
            # Taille des fichiers après conversion
            $taillePPTX += (Get-Item $newPath).Length
            $nbConvertis++

            # Remettre les métadonnées sur les fichiers .pptx
            $newFile = Get-Item $newPath
            $newFile.CreationTime = $creationTime
            $newFile.LastWriteTime = $lastWriteTime
            $newFile.LastAccessTime = $lastAccessTime

            $acl = Get-Acl $newFile.FullName
            $acl.SetOwner([System.Security.Principal.NTAccount] $owner)
            Set-Acl $newFile.FullName $acl

            Write-Host "Conversion réussie : $($_.Name)"
        }
        else {
            Write-Warning "Conversion échouée : $($_.Name)"
            $nbFails++
        }
    }
    catch {
        Write-Warning "Erreur avec $($_.Name) : $($_.Exception.Message)"
        $nbFails++
    }

    # Backup du .ppt dans le dossiers "backup_ppt" au même endroit que les fichiers convertis
    $oldFolder = $path.Substring(0, $path.LastIndexOf("\")) + "\backup_ppt"
    if (-not (Test-Path $oldFolder)) {
        New-Item $oldFolder -Type Directory | Out-Null
    }
    Move-Item $_.FullName $oldFolder
}

# Fermer PowerPoint
$ppt.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
$ppt = $null
[gc]::Collect()
[gc]::WaitForPendingFinalizers()

# Définir le chemin du fichier log
$logFile = "your_path\convert_ppt_log.txt"

# Résumé visuel dans la fenetre PS
Write-Host "`n=== Résumé conversion ==="
Write-Host "Dossier traité          : $folderpath"
Write-Host "Fichiers convertis      : $nbConvertis"
Write-Host "Fichiers en échec       : $nbFails"
Write-Host "Total taille PPT avant  : " ([Math]::Round($taillePPT / 1MB, 2)) "MB"
Write-Host "Total taille PPTX après : " ([Math]::Round($taillePPTX / 1MB, 2)) "MB"
Write-Host "Gain                    : " ([Math]::Round(($taillePPT - $taillePPTX) / 1MB, 2)) "MB"

if ($taillePPT -gt 0) {
    $pourcentageGain = (($taillePPT - $taillePPTX) / $taillePPT) * 100
    Write-Host "Réduction               : " ([Math]::Round($pourcentageGain, 2)) "%"
}
Write-Host ""

# Contenu du fichier de log
$logContent = @"
[$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")]
Dossier traité          : $folderpath
Fichiers convertis      : $nbConvertis
Fichiers en échec       : $nbFails
Total taille PPT avant  : $([Math]::Round($taillePPT / 1MB, 2)) MB
Total taille PPTX après : $([Math]::Round($taillePPTX / 1MB, 2)) MB
Gain                    : $([Math]::Round(($taillePPT - $taillePPTX) / 1MB, 2)) MB
Reduction               : $([Math]::Round($pourcentageGain, 2)) %
------------------------------------------------------------
"@

$logContent | Out-File -FilePath $logFile -Append -Encoding utf8

# Confirmation pour suppression des backup, si "Oui" alors suppresion puis fin de script, si "Non" alors fin du script
$title = 'Confirmation'
$question = 'Après vérification, souhaitez-vous supprimer les dossiers de backup .ppt ?
     Si vous choisissez [Non], il faudra les supprimer manuellement'
$choices = '  &Oui  ', '  &Non  '

$decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
if ($decision -eq 0) {
    Get-Childitem -Path $folderpath -Include backup_ppt -Recurse -Force | Remove-Item -Force -Recurse
    Write-Host "Les dossiers backup ont été supprimés"
}
else {
    Write-Host 'Fin de tâche, clôture du script'
    Start-Sleep -Seconds 2
}

# DOC to DOCX Batch convert script
# Inspiré de la version XLSX
# Édité by Logan

$folderpath = "your_path"
$filetype   = "*.doc"

# Chargement Word Interop, A commenté pour desactivé suivant la version d'excel installé
Add-Type -AssemblyName Microsoft.Office.Interop.Word

# Constante format DOCX
$wdFormatXMLDocument = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatXMLDocument

# Lance Word en mode invisible
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0  # wdAlertsNone

# Remise à zéro des compteurs de poids de fichiers, de nombre de fichiers convertis et d'erreurs
$tailleDOC  = 0
$tailleDOCX = 0
$nbConvertis = 0
$nbFails = 0

# Conversion des .doc en .docx
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
        $tailleDOC += $_.Length

        # Ouvrir et convertir
        $doc = $word.Documents.Open($_.FullName, $false, $true)
        $newPath = $path + ".docx"
        $doc.SaveAs([ref] $newPath, [ref] $wdFormatXMLDocument)
        $doc.Close()

        if (Test-Path $newPath) {
            # Taille des fichiers après conversion
            $tailleDOCX += (Get-Item $newPath).Length
            $nbConvertis++

            # Remettre les métadonnées sur les fichiers .docx
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

    # Backup du .doc dans le dossiers "backup_doc" au même endroit que les fichiers convertis
    $oldFolder = $path.Substring(0, $path.LastIndexOf("\")) + "\backup_doc"
    if (-not (Test-Path $oldFolder)) {
        New-Item $oldFolder -Type Directory | Out-Null
    }
    Move-Item $_.FullName $oldFolder
}

# Fermer Word
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
$word = $null
[gc]::Collect()
[gc]::WaitForPendingFinalizers()

# Définir le chemin du fichier log
$logFile = "your_path\convert_doc_log.txt"

# Résumé visuel dans la fenetre PS
Write-Host "`n=== Résumé conversion ==="
Write-Host "Dossier traité          : $folderpath"
Write-Host "Fichiers convertis      : $nbConvertis"
Write-Host "Fichiers en échec       : $nbFails"
Write-Host "Total taille DOC avant  : " ([Math]::Round($tailleDOC / 1MB, 2)) "MB"
Write-Host "Total taille DOCX après : " ([Math]::Round($tailleDOCX / 1MB, 2)) "MB"
Write-Host "Gain                    : " ([Math]::Round(($tailleDOC - $tailleDOCX) / 1MB, 2)) "MB"

if ($tailleDOC -gt 0) {
    $pourcentageGain = (($tailleDOC - $tailleDOCX) / $tailleDOC) * 100
    Write-Host "Réduction               : " ([Math]::Round($pourcentageGain, 2)) "%"
}
Write-Host ""

# Contenu du fichier de log
$logContent = @"
[$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")]
Dossier traité          : $folderpath
Fichiers convertis      : $nbConvertis
Fichiers en échec       : $nbFails
Total taille DOC avant  : $([Math]::Round($tailleDOC / 1MB, 2)) MB
Total taille DOCX après : $([Math]::Round($tailleDOCX / 1MB, 2)) MB
Gain                    : $([Math]::Round(($tailleDOC - $tailleDOCX) / 1MB, 2)) MB
Réduction               : $([Math]::Round($pourcentageGain, 2)) %
------------------------------------------------------------
"@

$logContent | Out-File -FilePath $logFile -Append -Encoding utf8

# Confirmation pour suppression des backup, si "Oui" alors suppresion puis fin de script, si "Non" alors fin du script
$title = 'Confirmation'
$question = 'Après vérification, souhaitez-vous supprimer les dossiers de backup .doc ?
     Si vous choisissez [Non], il faudra les supprimer manuellement'
$choices = '  &Oui  ', '  &Non  '

$decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
if ($decision -eq 0) {
    Get-Childitem -Path $folderpath -Include backup_doc -Recurse -Force | Remove-Item -Force -Recurse
    Write-Host "Les dossiers backup ont été supprimés"
}
else {
    Write-Host 'Fin de tâche, clôture du script'
    Start-Sleep -Seconds 2
}

# XLS to XLSX Batch convert script
# Forked from https://gist.github.com/gabceb/954418 
# Édité by Logan


#Chemin et extension de travail
$folderpath = "your_path"
$filetype   = "*.xls"

# Chargement Excel Interop, A commenté pour desactivé suivant la version d'excel installé
Add-Type -AssemblyName Microsoft.Office.Interop.Excel

# Constante format XLSX
$xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault

# Lance Excel en mode invisible
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# Remise à zéro des compteurs de poids de fichiers, de nombre de fichiers convertis et d'erreurs
$tailleXLS  = 0
$tailleXLSX = 0
$nbConvertis = 0
$nbFails = 0

# Conversion des .xls en .xlsx
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
        $tailleXLS += $_.Length

        # Ouvrir et convertir
        $workbook = $excel.Workbooks.Open($_.FullName)
        $newPath = $path + ".xlsx"
        $workbook.SaveAs($newPath, $xlFixedFormat)
        $workbook.Close()

        if (Test-Path $newPath) {
            # Taille des fichiers après conversion
            $tailleXLSX += (Get-Item $newPath).Length
            $nbConvertis++

            # Remettre les métadonnées sur les fichiers .xlsx
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

    # Backup du .xls dans le dossiers "backup_xls" au même endroit que les fichiers convertis
    $oldFolder = $path.Substring(0, $path.LastIndexOf("\")) + "\backup_xls"
    if (-not (Test-Path $oldFolder)) {
        New-Item $oldFolder -Type Directory | Out-Null
    }
    Move-Item $_.FullName $oldFolder
}

# Fermer Excel
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
$excel = $null
[gc]::Collect()
[gc]::WaitForPendingFinalizers()

# Définir le chemin du fichier log
$logFile = "your_path\convert_xls_log.txt"

# Résumé visuel dans la fenetre PS
Write-Host "`n=== Résumé conversion ==="
Write-Host "Dossier traité          : $folderpath"
Write-Host "Fichiers convertis      : $nbConvertis"
Write-Host "Fichiers en échec       : $nbFails"
Write-Host "Total taille XLS avant  : " ([Math]::Round($tailleXLS / 1MB, 2)) "MB"
Write-Host "Total taille XLSX après : " ([Math]::Round($tailleXLSX / 1MB, 2)) "MB"
Write-Host "Gain                    : " ([Math]::Round(($tailleXLS - $tailleXLSX) / 1MB, 2)) "MB"

if ($tailleXLS -gt 0) {
    $pourcentageGain = (($tailleXLS - $tailleXLSX) / $tailleXLS) * 100
    Write-Host "Réduction               : " ([Math]::Round($pourcentageGain, 2)) "%"
}
Write-Host ""

# Contenu du fichier de log
$logContent = @"
[$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")]
Dossier traité          : $folderpath
Fichiers convertis      : $nbConvertis
Fichiers en échec       : $nbFails
Total taille XLS avant  : $([Math]::Round($tailleXLS / 1MB, 2)) MB
Total taille XLSX après : $([Math]::Round($tailleXLSX / 1MB, 2)) MB
Gain                    : $([Math]::Round(($tailleXLS - $tailleXLSX) / 1MB, 2)) MB
Réduction               : $([Math]::Round($pourcentageGain, 2)) %
------------------------------------------------------------
"@

$logContent | Out-File -FilePath $logFile -Append -Encoding utf8

# Confirmation pour suppression des backup, si "Oui" alors suppresion puis fin de script, si "Non" alors fin du script
$title = 'Confirmation'
$question = 'Après vérification, souhaitez-vous supprimer les dossiers de backup .xls ?
     Si vous choisissez [Non], il faudra les supprimer manuellement'
$choices = '  &Oui  ', '  &Non  '

$decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
if ($decision -eq 0) {
    Get-ChildItem -Path $folderpath -Include backup_xls -Recurse -Force | Remove-Item -Force -Recurse
    Write-Host "Les dossiers backup ont été supprimés"
}
else {
    Write-Host 'Fin de tâche, clôture du script'
    Start-Sleep -Seconds 2
}

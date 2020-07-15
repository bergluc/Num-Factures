#Obtention de la date et de l'heure
$TimeStamp = Get-Date -Format yyyyMMdd_HHmmss
$An = Get-Date -Format yyyy

#Ajout de la date/heure dans le fichier de transactions (log)
Add-Content -Path 'E:\Fichiers\Rapports Regionaux\Factures envoyées au siège social\Numérisation\transactions.txt' -Value $TimeStamp

#Création du répertoire pour l'historique des fichiers traités
$newdir = 'E:\Fichiers\Rapports Regionaux\Factures envoyées au siège social\Numérisation\historique\' + $TimeStamp 
New-Item -Path $newdir -ItemType Directory

#À changer par move-item en production --> faut alors enlever le -Recurse?
copy-item \\mfp11222796\FILE_SHARE\002-Facturation-Numériser_Factures\* -Recurse -destination E:\Fichiers\'Rapports Regionaux'\'Factures envoyées au siège social'\Numérisation\AT
$Source = '300'

#Enlever le sous-dossier TestNum en production
$destdir = 'E:\Fichiers\Rapports Regionaux\Factures envoyées au siège social\TestNum\' + $An + '\' + $Source

#Changement du répertoire de traitement
Set-Location 'E:\Fichiers\Rapports Regionaux\Factures envoyées au siège social\Numérisation\AT'

#Traitement OCR des factures numérisées, journalisation et transfert des fichiers TIF dans l'historique
Get-ChildItem -Filter '*.tif' -Recurse | ForEach-Object {
    $newname =  $Source + '_' + $TimeStamp + '_' + $_.BaseName
    C:\'Program Files'\Tesseract-OCR\tesseract.exe $_.FullName $newname -l fra pdf
	Add-Content -Path 'E:\Fichiers\Rapports Regionaux\Factures envoyées au siège social\Numérisation\transactions.txt' -Value $_.Name
	move-item -path $_.FullName -destination $newdir 
}

#Transfert des factures PDF dans le répertoire de recherche de la comptabilité
Get-ChildItem -Filter '*.pdf' | ForEach-Object {
    move-item -path $_.FullName -destination $destdir
}
#Suppressions des répertoires après les déplacements de fichiers.
remove-item -path .\*

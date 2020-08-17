#Obtention de la date et de l'heure
$TimeStamp = Get-Date -Format yyyyMMdd_HHmmss

#Ajout de la date/heure dans le fichier de transactions (log)
Add-Content -Path 'C:\temp\transactions.txt' -Value $TimeStamp

#Création du répertoire pour l'historique des fichiers traités
$newdir = 'C:\temp\historique\' + $TimeStamp 
New-Item -Path $newdir -ItemType Directory

#À changer par move-item en production --> faut alors enlever le -Recurse?
copy-item \\mfp11222796\FILE_SHARE\002-Facturation-Numériser_Factures\* -Recurse -destination C:\temp\AT

#Changement du répertoire de traitement
Set-Location 'C:\temp\AT'

#Traitement OCR des factures numérisées, journalisation et transfert des fichiers TIF dans l'historique
Get-ChildItem -Filter '*.tif' -Recurse | ForEach-Object {
    $newname = $_.BaseName + $TimeStamp
    C:\'Program Files'\Tesseract-OCR\tesseract.exe $_.FullName $newname -l fra pdf
	Add-Content -Path 'C:\temp\transactions.txt' -Value $_.Name
	move-item -path $_.FullName -destination $newdir 
}

#Transfert des factures PDF dans le répertoire de recherche de la comptabilité
Get-ChildItem -Filter '*.pdf' | ForEach-Object {
    move-item -path $_.FullName -destination C:\Users\lroberge\Desktop\Index
}
#Suppressions des répertoires après les déplacements de fichiers.
remove-item -path .\*
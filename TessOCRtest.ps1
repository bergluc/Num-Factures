#Adresse du dossier partagé sur l'imprimante où sont placés les documents numérisés suite à la sélection du modèle créé pour la facturation
Set-Location '\\MFP11222796\FILE_SHARE\002-Facturation-Numériser_Factures'

#Obtention de la date et de l'heure
$TimeStamp = Get-Date -Format yyyyMMdd_HHmmss

#Ajout de la date/heure dans le fichier de transactions (log)
Add-Content -Path 'C:\temp\transactions.txt' -Value $TimeStamp

#Création du répertoire pour l'historique des fichiers traités
$newdir = 'historique\' + $TimeStamp 
New-Item -Path $newdir -ItemType Directory

#Traitement OCR des factures numérisées, journalisation et transfert des fichiers TIF dans l'historique
Get-ChildItem -Filter '*.tif' -Recurse | ForEach-Object {
    $newname = $_.BaseName + $TimeStamp
    C:\Users\lroberge\AppData\Local\Tesseract-OCR\tesseract.exe $_.Name $newname -l fra pdf
	Add-Content -Path 'C:\temp\transactions.txt' -Value $_.Name
	move-item -path $_.Name -destination $newdir 
}

#Transfert des factures PDF dans le répertoire de recherche de la comptabilité
Get-ChildItem -Filter '*.pdf' | ForEach-Object {
    move-item -path $_.Name -destination C:\Users\lroberge\Desktop\Index
}

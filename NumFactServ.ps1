#Obtention de la date et de l'heure
$TimeStamp = Get-Date -Format yyyyMMdd_HHmmss
$An = Get-Date -Format yyyy
$Fact = 'E:\Fichiers\Rapports Regionaux\Factures envoyées au siège social\Numérisation\'
$Mod = '002-Facturation-Numériser_Factures\'

#Ajout de la date/heure dans le fichier de transactions (log)
Add-Content -Path $Fact + 'transactions.txt' -Value $TimeStamp

#Futur Import=Csv et boucle 
$Inst = '300'
$Doss = '\\mfp11222796\FILE_SHARE\'
$Impr = 'PHOACHAT1'

#Création du répertoire pour l'historique des fichiers traités
$newdir = $Fact + 'historique\' + $TimeStamp +'_' + $Inst
New-Item -Path $newdir -ItemType Directory

#À changer par move-item en production --> faut alors enlever le -Recurse?
$Source = $Doss + $Mod + '*'
copy-item $Sources -Recurse -destination $Fact + 'AT'

#Enlever le sous-dossier TestNum en production
$destdir = $Fact + 'TestNum\' + $An + '\' + $Inst

#Changement du répertoire de traitement
Set-Location $Fact + 'AT'

#Traitement OCR des factures numérisées, journalisation et transfert des fichiers TIF dans l'historique
Get-ChildItem -Filter '*.tif' -Recurse | ForEach-Object {
    $newname =  $Inst + '_' + $Impr + '_' + $TimeStamp + '_' + $_.BaseName
    C:\'Program Files'\Tesseract-OCR\tesseract.exe $_.FullName $newname -l fra pdf
	Add-Content -Path $Fact + 'transactions.txt' -Value $_.Name
	move-item -path $_.FullName -destination $newdir 
}

#Transfert des factures PDF dans le répertoire de recherche de la comptabilité
Get-ChildItem -Filter '*.pdf' | ForEach-Object {
    move-item -path $_.FullName -destination $destdir
}
#Suppressions des répertoires après les déplacements de fichiers.
remove-item -path .\*

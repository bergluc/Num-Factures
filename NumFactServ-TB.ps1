#Obtention de la date et de l'heure

$TimeStamp = Get-Date -Format yyyyMMdd_HHmmssff

$An = Get-Date -Format yyyy

$Fact = '\\cldvsrvfs01\fichiers\Rapports Regionaux\Factures envoyées au siège social\Numérisation\'

$Mod = '002-Facturation-Numériser_Factures\'



#Ajout de la date/heure dans le fichier de transactions (log)

$Trans = $Fact + 'transactions.txt'
Add-Content -Path $Trans -Value $TimeStamp



#Import=Csv et boucle 

$FilePath = "Imprimantes.csv"

$Contenu = Import-CSV $FilePath



ForEach ($Imprimante in $Contenu) {

$Impr = $($Imprimante.Impr)

$Inst = $($Imprimante.Inst)

$Dossier = $($Imprimante.Dossier)



#Création du répertoire pour l'historique des fichiers traités

#$newdir = $Fact + 'historique\' + $TimeStamp +'_' + $Inst

#New-Item -Path $newdir -ItemType Directory



#À changer par move-item en production --> faut alors enlever le -Recurse?

$Source = $Dossier + $Mod + '*'
$SourDIR = $Dossier + $Mod
$Message = "Je tente de copier à partir de " + $Source
Add-Content -Path $Trans -Value $Message
$AT = $Fact + 'AT'
If(!(Test-Path $AT)){New-Item -Path $AT -ItemType Directory}
If(Test-Path $SourDIR){
	copy-item $Source -Recurse -destination $AT


#Enlever le sous-dossier TestNum en production

#$destdir = $Fact + 'TestNum\' + $An + '\' + $Inst



#Changement du répertoire de traitement

Set-Location $AT



#Traitement OCR des factures numérisées, journalisation et transfert des fichiers TIF dans l'historique

Get-ChildItem -Filter '*.tif' -Recurse | ForEach-Object {

     $TimeStamp = Get-Date -Format yyyyMMdd_HHmmssff
     $newname =  $Inst + '_' + $Impr + '_' + $TimeStamp + '_' + $_.BaseName

#    C:\'Program Files'\Tesseract-OCR\tesseract.exe $_.FullName $newname -l fra pdf

	Add-Content -Path $Trans -Value $newname
#	move-item -path $_.FullName -destination $newdir 

}



#Transfert des factures PDF dans le répertoire de recherche de la comptabilité

#Get-ChildItem -Filter '*.pdf' | ForEach-Object {

#    move-item -path $_.FullName -destination $destdir

#}

#Suppressions des répertoires après les déplacements de fichiers.

#remove-item -path .\*

}
}
#Obtention de l'année de la date et de l'heure
$TimeStamp = Get-Date -Format yyyyMMdd_HHmmssff
$An = Get-Date -Format yyyy

#Dossier des factures et dossiers de travail
$Fact = '\\cldvsrvfs01\fichiers\Rapports Regionaux\Factures envoyées au siège social\'
$FactDT = $Fact + 'Numérisation\'
If(!(Test-Path $FactDT)){New-Item -Path $FactDT -ItemType Directory}
$Hist = $FactDT + 'historique\'
If(!(Test-Path $Hist)){New-Item -Path $Hist -ItemType Directory}
$AT = $FactDT + 'AT'
If(!(Test-Path $AT)){New-Item -Path $AT -ItemType Directory}

#Dossier des factures numérisées en fonction du nom de modèle sur les copieurs
$Mod = '002-Facturation-Numériser_Factures\'

#Ajout de la date/heure dans le fichier de transactions (log)
$Trans = $FactDT + 'transactions.txt'
Add-Content -Path $Trans -Value $TimeStamp

#Importation des imprimantes
$FilePath = 'Imprimantes.csv'
If(!(Test-Path $FilePath)){
	Add-Content -Path $Trans -Value 'Pas de fichier Imprimantes.csv ---> Fin de traitement'
	Exit
}
$Contenu = Import-CSV $FilePath

#Changement du répertoire de traitement
Set-Location $AT

#Traitement des factures sur chaque imprimamte
ForEach ($Imprimante in $Contenu) {
	$Impr = $($Imprimante.Impr)
	$Inst = $($Imprimante.Inst)
	$Dossier = $($Imprimante.Dossier)
	$Copieur = 'Imprimante : ' + $Impr + ' - ' + $Dossier
	Add-Content -Path $Trans -Value $Copieur

	#Initialisation des dossiers pour la copie des factures numérisées
	$Source = $Dossier + $Mod + '*'
	$SourDIR = $Dossier + $Mod

	#Traitement des factures seulement s'il y en a dans le dossier de l'imprimante
	If(Test-Path $SourDIR){

		#Création du répertoire pour l'historique des fichiers traités
		$newdir = $Hist + $TimeStamp + '_' + $Inst + '\'
		New-Item -Path $newdir -ItemType Directory

		#Copie des factures dans le dossier AT (à traiter)
		#et suppression des factures sur le photocopieur
		copy-item $Source -Recurse -destination $AT
    remove-item $Source -Recurse

		#Définir le dossier de destination des factures traitées
		# ---> Enlever le sous-dossier TestNum en production
		# $destdir = $Fact + 'TestNum\' + $An + '\' + $Inst
		$destdir = $Fact + $An + '\' + $Inst
		If(!(Test-Path $destdir)){New-Item -Path $destdir -ItemType Directory}

		#Traitement OCR des factures numérisées, journalisation et transfert des fichiers TIF dans l'historique
		Get-ChildItem -Filter '*.tif' -Recurse | ForEach-Object {
			$TimeStamp = Get-Date -Format yyyyMMdd_HHmmssff
	    		$newname =  $Inst + '_' + $Impr + '_' + $TimeStamp + '_' + $_.BaseName
	    		C:\'Program Files'\Tesseract-OCR\tesseract.exe $_.FullName $newname -l fra pdf
			$Journal = $_.Name + " ---> " + $newname
			Add-Content -Path $Trans -Value $Journal
			$NouvDossNom = $newdir + $TimeStamp + $_.Name
			move-item -path $_.FullName -destination $NouvDossNom
		}

		#Transfert des factures PDF dans le répertoire de recherche de la comptabilité
		Get-ChildItem -Filter '*.pdf' | ForEach-Object {
	    		move-item -path $_.FullName -destination $destdir
		}

		#Suppressions des répertoires après les déplacements de fichiers.
		remove-item -path .\* -Force
	}
}

#Obtention de la date et de l'heure
$TimeStamp = Get-Date -Format yyyyMMdd_HHmmssff
$An = Get-Date -Format yyyy
$Fact = '\\cldvsrvfs01\fichiers\Rapports Regionaux\Factures envoyées au siège social\Numérisation\'
$Mod = '002-Facturation-Numériser_Factures\'

#Ajout de la date/heure dans le fichier de transactions (log)
$Trans = $Fact + 'transactions.txt'
Add-Content -Path $Trans -Value $TimeStamp

#Importation des imprimantes
# --- > Ajouter un test : si le fichier n'existe pas, arrêter le script
$FilePath = "Imprimantes.csv"
$Contenu = Import-CSV $FilePath

#Traitement des factures sur chaque imprimamte
ForEach ($Imprimante in $Contenu) {
	$Impr = $($Imprimante.Impr)
	$Inst = $($Imprimante.Inst)
	$Dossier = $($Imprimante.Dossier)

	#Création du dossier AT s'il n'existe pas
	$AT = $Fact + 'AT'
	If(!(Test-Path $AT)){New-Item -Path $AT -ItemType Directory}

	#Initialisation des dossiers pour la copie des factures numérisées
	$Source = $Dossier + $Mod + '*'
	$SourDIR = $Dossier + $Mod

	#Traitement des factures seulement s'il y en a dans le dossier de l'imprimante
	If(Test-Path $SourDIR){

		#Création du répertoire pour l'historique des fichiers traités
		$newdir = $Fact + 'historique\' + $TimeStamp +'_' + $Inst
		New-Item -Path $newdir -ItemType Directory

		#Copie des factures dans le dossier AT (à traiter)
		# ---> À changer par move-item en production --> faut alors enlever le -Recurse	
		copy-item $Source -Recurse -destination $AT

		#Définir le dossier de destination des factures traitées
		# ---> Enlever le sous-dossier TestNum en production
		$destdir = $Fact + 'TestNum\' + $An + '\' + $Inst

		#Changement du répertoire de traitement
		Set-Location $AT

		#Traitement OCR des factures numérisées, journalisation et transfert des fichiers TIF dans l'historique
		Get-ChildItem -Filter '*.tif' -Recurse | ForEach-Object {
	    		$newname =  $Inst + '_' + $Impr + '_' + $TimeStamp + '_' + $_.BaseName
	    		C:\'Program Files'\Tesseract-OCR\tesseract.exe $_.FullName $newname -l fra pdf
			Add-Content -Path $Trans -Value $_.Name
			move-item -path $_.FullName -destination $newdir 
		}

		#Transfert des factures PDF dans le répertoire de recherche de la comptabilité
		Get-ChildItem -Filter '*.pdf' | ForEach-Object {
	    		move-item -path $_.FullName -destination $destdir
		}
	
		#Suppressions des répertoires après les déplacements de fichiers.
		remove-item -path .\*
	}
}

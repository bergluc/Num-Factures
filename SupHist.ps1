#Dossier des factures et dossiers de travail
$Fact = '\\cldvsrvfs01\fichiers\Rapports Regionaux\Factures envoyées au siège social\'
$FactDT = $Fact + 'Numérisation\'
$Hist = $FactDT + 'historique\'
$Log = $FactDT+ 'SuppressionsHistorique.txt'

$Folder = $Hist

#Delete files older than 6 months
Get-ChildItem $Folder -Recurse -Force -ea 0 |
? {!$_.PsIsContainer -and $_.LastWriteTime -lt (Get-Date).AddDays(-180)} |
ForEach-Object {
#   $_ | del -Force
   $_.FullName | Out-File $Log -Append
}

#Delete empty folders and subfolders
Get-ChildItem $Folder -Recurse -Force -ea 0 |
? {$_.PsIsContainer -eq $True} |
? {$_.getfiles().count -eq 0} |
ForEach-Object {
#    $_ | del -Force
    $_.FullName | Out-File $log -Append
}

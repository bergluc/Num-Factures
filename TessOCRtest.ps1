Set-Location 'C:\Users\lroberge\Desktop\FAT'

$TimeStamp = Get-Date -Format yyyyMMdd_HHmmss
Add-Content -Path 'Transactions.txt' -Value $TimeStamp
$newdir = 'Historique\' + $TimeStamp 
New-Item -Path $newdir -ItemType Directory

Get-ChildItem -Filter '*.tif' -Recurse | ForEach-Object {
    $newname = $_.BaseName + $TimeStamp
    C:\Users\lroberge\AppData\Local\Tesseract-OCR\tesseract.exe $_.Name $newname -l fra pdf
	Add-Content -Path 'C:\Users\lroberge\Desktop\FAT\Transactions.txt' -Value $_.Name
	move-item -path $_.Name -destination $newdir 
}

Get-ChildItem -Filter '*.pdf' | ForEach-Object {
    move-item -path $_.Name -destination C:\Users\lroberge\Desktop\Index
}

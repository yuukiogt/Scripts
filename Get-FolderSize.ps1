[string]$targetFolder = Read-Host "FolderName "

$folders = Get-ChildItem -Path $targetFolder -Recurse | Where-Object PSIsContainer
[array]$volume = foreach ($folder in $folders) {
    $subFolderItems = (Get-ChildItem $folder.FullName | Where-Object Length | Measure-Object Length -Sum)
    [PSCustomObject]@{
        Fullname = $folder.FullName
        MB=[decimal]("{0:N2}" -f ($subFolderItems.Sum / 1MB))
    }
}

$volume | Sort-Object MB -Descending | Out-GridView
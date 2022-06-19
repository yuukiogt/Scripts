$rootPath = ".\Path"

$childFoldersMap = @{}
$childFolders = Get-ChildItem $rootPath -Recurse -Directory | Select-Object FullName
foreach($childFolder in $childFolders) {
    $childFoldersMap[$childFolder.FullName] = ($childFolder.FullName.Split("\")).Length
}

$childFoldersMap = $childFoldersMap.GetEnumerator() | Sort-Object -Property Value -Descending

foreach ($folder in $childFoldersMap.GetEnumerator()) {
    $tmpItems = Get-ChildItem $folder.Key | Sort-Object LastWriteTime -Descending
    if($tmpItems.Length -eq 0) {
        continue
    }
    $latestItemWriteTime = $tmpItems[0].LastWriteTime
    Set-ItemProperty -Path $folder.Key -Name Attributes -Value "Normal"
    Set-ItemProperty -Path $folder.Key -Name LastWriteTime -Value $latestItemWriteTime
}

$rootChildItems = Get-ChildItem $rootPath | Sort-Object LastWriteTime -Descending
$latestItemWriteTime = $rootChildItems[0].LastWriteTime
Set-ItemProperty -Path $rootPath -Name LastWriteTime -Value $latestItemWriteTime
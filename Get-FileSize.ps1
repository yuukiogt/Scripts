
function GetFileSize($parent_folder, $count = 100) {

write-progress "進捗表示のために総ファイル数を確認しています....";
$totalfilecount=(Get-ChildItem -Path $parent_folder -Recurse |measure).count

$cnt=0

Get-ChildItem -Path $parent_folder -Recurse |
    Select-Object Length,fullname,LastWriteTime |
        %{New-Object psobject -Property @{SizeMB=[math]::round($_.Length/1024/1024);Filename=$_.fullname;LastWrite=$_.LastWriteTime};$cnt++;Write-Progress "現在の処理ファイル数：$cnt/$totalfilecount" -PercentComplete ($cnt/$totalfilecount*100)} |
            Sort-Object -Descending SizeMB |
                Select-Object -first 10 SizeMB,Filename,LastWrite

}

do {
    $path = Read-Host "Folder Path"
} until((Test-Path $path))

$count = Read-Host "Count"

GetFileSize($path, $count)
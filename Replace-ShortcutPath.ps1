
while($True) {
    $location = Read-Host "Path "

    if(Test-Path $location) {
        break;
    }
}

Set-Location $location

$wsh = New-Object -ComObject wscript.shell
$links = Get-ChildItem -Recurse -Include *.lnk

[string]$searchString = Read-Host "検索する文字列 "
[string]$replaceString = Read-Host "置換後の文字列 "

foreach ($link in $links) {
    $sht = $wsh.createshortcut($link)
    $fullName = $sht.fullname
    [string]$linkName = $sht.targetpath
    $newPath = $linkName.Replace($searchString, $replaceString)
    $sht = $wsh.createshortcut($fullName)
    $sht.targetpath = $newPath
    $sht.save()
}
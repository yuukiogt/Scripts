$base = "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
$wow64 = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
$path = @(("HKLM:" + $base), ("HKCU:" + $base))
if(Test-Path $wow64){
    $path += $wow64
}

$apps = Get-ChildItem -Path $path |
ForEach-Object {Get-ItemProperty -Path $_.PsPath} |
Where-Object {$_.systemcomponent -ne 1} |
Where-Object {$_.parentkeyname -eq $null} |
Where-Object {$_.DisplayName -ne $null} |
Select-Object DisplayName, Publisher, Version |
Sort-Object DisplayName

$apps | Out-GridView
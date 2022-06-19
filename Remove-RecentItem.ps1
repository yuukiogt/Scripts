filter Get-Shortcut() {
    $shell = new-object -comobject WScript.Shell
    return $shell.CreateShortcut($_)
}

$profiles = Get-ChildItem "C:\Users"
foreach ($profile in $profiles) {
    $targetLink = $profile.FullName + "target.lnk"
    $isExistlnk = Test-Path $targetLink
    if ($isExistlnk) {
        $t = $targetLink | Get-Shortcut | Select-Object -property TargetPath
        Remove-Item -Force $t.TargetPath
    }
}
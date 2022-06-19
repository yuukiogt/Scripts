
$PacksParentDirectory = "C:\Windows\diagnostics\system"
$Packs = Get-ChildItem $PacksParentDirectory | Where-Object { $_.PSIsContainer }

$Index = 0
$PackNames = @()
$PackNamesString = $Packs |
ForEach-Object {
    $PackNames += $_.Name
    "$($index): " + $_.Name
    $index++
}

$PackNamesString += "$($index): All"
$PackNamesString += "q: quit"

while (1) {
    $Target = Read-Host @"
    $PackNamesString
"@

    if ($Target -eq 'q') {
        break
    }

    if ($Target -eq $PackNamesString.Count - 2) {
        foreach ($Name in $Packs.Name) {
            $Path = Join-Path $PacksParentDirectory $Name
            Write-Host -f Cyan $Name
            Get-TroubleshootingPack -Path $Path -Verbose | Invoke-TroubleshootingPack -Verbose
        }
    }
    else {
        $Path = Join-Path $PacksParentDirectory $PackNames[$Target]
        Write-Host -f Cyan $PackNames[$Target]
        Get-TroubleshootingPack -Path $Path -Verbose | Invoke-TroubleshootingPack -Verbose
    }
}
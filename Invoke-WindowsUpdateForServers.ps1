try {
    Set-Location $PSScriptRoot

    $servers = Get-ADComputer -Filter { OperatingSystem -like "*Server*" -and Enabled -eq $True }
    $time = Get-Date -Format yyyyMMddhhmm

    $pass = Read-Host "pass:"
    foreach ($server in $servers) {
        Start-Job { powershell -File "$($PSScriptRoot)\WindowsUpdateForServers.ps1" $server.Name $pass } -Name "$($time)_$($server.Name)"
    }
}
catch {
    Write-Host $_.Exception.Message
}
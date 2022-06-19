
$servers = Get-ADComputer -Filter { OperatingSystem -like "*Server*" -and Enabled -eq $True } |
Where-Object { $_.DistinguishedName -like "*OU=Server*" -or $_.DistinguishedName -like "*OU=Domain*" }

$index = 0
$serverNames = @()
$serverNamesString = $servers |
ForEach-Object {
    $serverNames += $_.Name
    "$($index): " + $_.Name
    $index++
}

$serverNamesString += "$($index): All"
$serverNamesString += "q: quit"

$credential = Get-Credential

function SecureString2PlainString($SecureString) {
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
    $PlainString = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($BSTR)

    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)

    return $PlainString
}

function InvokeTaskScheduleCommand($serverName) {
    Invoke-Command -ComputerName $serverName -Credential $credential -ScriptBlock {
        $updateSession = New-Object -com Microsoft.Update.Session
        $searcher = $updateSession.CreateUpdateSearcher()
        $searchResult = $searcher.search("IsInstalled=0 and Type='software'")
        $searchResult.Updates | ForEach-Object { $_.title -replace ".*(KB\d+).*", "`$1`t$&" }
        $updatesToDownload = New-Object -com Microsoft.Update.UpdateColl
        $searchResult.Updates | Where-Object { -not $_.InstallationBehavior.CanRequestUserInput } | Where-Object { $_.EulaAccepted } | ForEach-Object { [void]$updatesToDownload.add($_) }
        $downloader = $updateSession.CreateUpdateDownloader()
        $downloader.Updates = $updatesToDownload
        $downloader.Download()
        $updatesToInstall = New-Object -com Microsoft.Update.UpdateColl
        $searchResult.Updates | Where-Object { $_.IsDownloaded } | ForEach-Object { [void]$updatesToInstall.add($_) }
        $installer = $updateSession.CreateUpdateInstaller()
        $installer.Updates = $updatesToInstall
        $installationResult = $installer.Install()
        $installationResult

        $taskTriggers = (Get-ScheduledTask | Where-Object TaskPath -eq \ | Where-Object TaskName -eq "Reboot").Triggers
        $prev = [DateTime]$taskTriggers.StartBoundary
        $now = Get-Date
        $next = $prev.AddDays(($now - $prev).Days + 1)
        if(($now - $prev).Days -eq 0) {
            $next.AddDays(-1)
        }

        schtasks /create /tn Reboot /tr "shutdown -r -t 60 -f" /sc Once /sd $next.ToString("yyyy/MM/dd") /st $next.ToString("HH:mm") /ru system /f

        Write-Host "Reboot タスクを $($next) にセットしました"
    } -ArgumentList $serverName
}

while (1) {
    $target = Read-Host @"
    $serverNamesString
"@

    if ($target -eq 'q') {
        break
    }

    if ($target -eq $serverNamesString.Count - 2) {
        $serversString = ""
        for($i = 0; $i -lt $servers.Length; $i++) {
            $serversString += $servers[$i].Name
            if($i -ne $servers.Length - 1) {
                $serversString += ", "
            }
        }
        InvokeTaskScheduleCommand $serversString
    }
    else {
        $targetName = $serverNames[$target]
        InvokeTaskScheduleCommand $targetName
    }
}
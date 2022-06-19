
$servers = Get-ADComputer -Filter { OperatingSystem -like "*Server*" -and Enabled -eq $True }

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
        $taskTriggers = (Get-ScheduledTask | Where-Object TaskPath -eq \ | Where-Object TaskName -eq "Reboot").Triggers
        $prev = [DateTime]$taskTriggers.StartBoundary
        $now = Get-Date
        $next = $prev.AddDays(($now - $prev).Days + 1)

        schtasks /create /tn Reboot /tr "shutdown -r -t 60 -f" /sc Once /sd $next.ToString("yyyy/MM/dd") /st $next.ToString("HH:mm") /ru system /f

        Write-Host "$($args[0]) の Reboot タスクを $($next) にセットしました"
    } -ArgumentList $serverName
}

while(1) {
    $target = Read-Host @"
    $serverNamesString
"@

    if($target -eq 'q') {
        break
    }

    if($target -eq $serverNamesString.Count - 2) {
        foreach ($server in $servers) {
            InvokeTaskScheduleCommand $server.Name
        }
    } else {
        $targetName = $serverNames[$target]
        InvokeTaskScheduleCommand $targetName
    }
}
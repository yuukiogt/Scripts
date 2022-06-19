$servers = Get-ADComputer -Filter { OperatingSystem -like "*Server*" -and Enabled -eq $True }

foreach ($server in $servers) {
    $server.Name
    Invoke-Command -ComputerName $server.Name -ScriptBlock {
        Get-ScheduledTaskInfo -TaskPath \Microsoft\Windows\Backup -TaskName Microsoft-Windows-WindowsBackup
    }
}
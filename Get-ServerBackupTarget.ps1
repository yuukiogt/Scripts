$servers = Get-ADComputer -Filter { OperatingSystem -like "*Server*" } |
Where-Object { $_.Enabled -eq $True } |
Select-Object Name

foreach ($server in $servers) {
    $connection = Test-Connection $server.Name -Quiet -Count 1
    if ($connection -eq $False) {
        Write-Host "$($server.Name) 接続できませんでした"
        continue;
    }

    $server.Name
    
    Invoke-Command -ComputerName $server.Name -ScriptBlock {
        $policy = Get-WBPolicy
        $target = Get-WBBackupTarget -Policy $policy

        Write-Host "---Get-WBBackupTarget---"
        $target

        Write-Host "---Get-WBJob---"
        Get-WBJob -Previous 1
    }
}
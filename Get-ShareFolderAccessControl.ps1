
$servers = Get-ADComputer -Filter { OperatingSystem -like "*Server*" -and Enabled -eq $True } | Select-Object Name

foreach($server in $servers) {
    Invoke-Command -ComputerName $server.Name -ScriptBlock {
        Get-SmbShare -Special:$False
    }
}
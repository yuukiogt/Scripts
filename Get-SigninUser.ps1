do {
    $HostName = Read-Host "Hostname"

    $isPingResp = Test-Connection -Quiet -ComputerName $HostName -Count 1
    if($isPingResp -eq $False){
        Write-Host "Ping応答無し"
        continue
    }

    $Query = "query user /server:${Hostname}"
    $Array = Invoke-Expression $Query

    $Array

    Write-Host "---`n"
} while ((Read-Host "quit (y/n)") -ne "y")
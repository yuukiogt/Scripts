$Servers = Get-ADComputer -Filter { OperatingSystem -like "*Server*" } |
Where-Object { $_.Enabled -eq $True }

$Result = @{ }
foreach ($Server in $Servers) {
    $NtfsError = Get-WinEvent -ComputerName $Server -MaxEvents 16 -FilterHashtable @{LogName = "System"; ID = 55 }
    if ($Null -ne $NtfsError) {
        $Result.Add($Server.Name, $NtfsError)
    }
}

$Result | Out-GridView
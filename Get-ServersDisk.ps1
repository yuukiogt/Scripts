$Servers = Get-ADComputer -Filter { OperatingSystem -like "*Server*" } |
Where-Object { $_.Enabled -eq $True }

$GetDiskResult = @{ }
foreach ($Server in $Servers) {
    $Disks = Get-Disk -CimSession $Server.Name
    if ($Null -ne $Disks) {
        $GetDiskResult.Add($Server.Name, $Disks)
    }
}

$GetDiskResult.GetEnumerator() |
Select-Object @{n = 'ServerName'; e = { $_.Name } },
@{n = 'Number'; e = { $_.Value.Number } },
@{n = 'FriendlyName'; e = { $_.Value.FriendlyName } },
@{n = 'SerialNumber'; e = { $_.Value.SerialNumber } },
@{n = 'HealthStatus'; e = { $_.Value.HealthStatus } },
@{n = 'OperationalStatus'; e = { $_.Value.OperationalStatus } },
@{n = 'TotalSize'; e = { $_.Value.Size } } | Out-GridView

$GetPhysicalDiskResult = @{ }
foreach ($Server in $Servers) {
    $PhysicalDisks = Get-PhysicalDisk -CimSession $Server.Name
    if ($Null -ne $PhysicalDisks) {
        $GetPhysicalDiskResult.Add($Server.Name, $PhysicalDisks)
    }
}

$GetPhysicalDiskResult.GetEnumerator() |
Select-Object @{n = 'ServerName'; e = { $_.Name } },
@{n = 'Number'; e = { $_.Value.Number } },
@{n = 'FriendlyName'; e = { $_.Value.FriendlyName } },
@{n = 'SerialNumber'; e = { $_.Value.SerialNumber } },
@{n = 'MediaType'; e = { $_.Value.MediaType } },
@{n = 'CanPool'; e = { $_.Value.CanPool } },
@{n = 'HealthStatus'; e = { $_.Value.HealthStatus } },
@{n = 'OperationalStatus'; e = { $_.Value.OperationalStatus } },
@{n = 'TotalSize'; e = { $_.Value.Size } } | Out-GridView
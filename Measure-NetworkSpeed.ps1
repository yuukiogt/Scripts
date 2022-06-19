[System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces() |
Where-Object { $_.OperationalStatus -eq "Up"} |
Select-Object Description,
@{label = "Speed(Mbps)"; expression = { $_.Speed / 1000 / 1000 } },
@{label = "Speed(Gbps)"; expression = { $_.Speed / 1000 / 1000 / 1000 }},
OperationalStatus |
Sort-Object Description |
Format-Table -AutoSize
$servers = Get-ADComputer -Properties IPv4Address -Filter { OperatingSystem -like "*Server*" } |
Where-Object { $_.Enabled -eq $True } |
Select-Object Name,IPv4Address

$ipString = [string]::Empty

$index = 0
$serverNames = @()
$serverIPs = @()
$serverNamesString = $servers |
ForEach-Object {
    $serverNames += $_.Name
    $serverIPs += $_.IPv4Address
    "$($index): " + $_.Name
    $index++
}

$serverNamesString += "$($index): All"
$serverNamesString += "q: quit"

while (1) {
    $target = Read-Host @"
    $serverNamesString
"@

    if ($target -eq 'q') {
        break
    }

    if ($target -eq $serverNamesString.Count - 2) {
        foreach ($name in $servers.IPv4Address) {
            $ipString += "$($name),"
        }

        $ipString = $ipString.Remove($ipString.LastIndexOf(','), 1)

        wmic /node:($ipString) logicaldisk where DriveType=3 get "FileSystem,FreeSpace,Name,Size,SystemName,VolumeName" /format:csv
    }
    else {
        wmic /node:($serverIPs[$target]) logicaldisk where DriveType=3 get "FileSystem,FreeSpace,Name,Size,SystemName,VolumeName" /format:list
    }
}
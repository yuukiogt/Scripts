$servers = Get-ADComputer -Properties IPv4Address -Filter { OperatingSystem -like "*Server*" } |
Where-Object { $_.Enabled -eq $True } |
Select-Object Name, IPv4Address

function GetWindowsServerBackupEvents([string]$serverName, [string]$eventID) {
    $query = "*[System[Provider[@Name='Microsoft-Windows-Backup'] and (EventID=${eventID})]]"
    $resultArray = wevtutil qe Microsoft-Windows-Backup /q:$query /r:$serverName /uni:true /f:text
    return $resultArray
}

function GetEvents([string]$taskName,[string]$serverName,[string]$eventID) {
    $query = "*[EventData[Data[@Name='TaskName']='${taskName}']] and *[System[(EventID=${eventID})]]"
    $resultArray = wevtutil qe Microsoft-Windows-TaskScheduler/Operational /q:$query /r:$serverName /uni:true /f:text
    return $resultArray
}

function SplitDateString($dateString) {
    [string]$result = $dateString.SubString($dateString.IndexOf(':') + 1, $dateString.Length - $dateString.IndexOf(':') - 1 ).Trim()
    return $result
}

function PushEvent($array) {
    for ($i = 3; $i -lt $array.Length; $i += 15) {
        $eventID = $i + 1
        if ($eventID -lt $array.Length) {
            $key = SplitDateString($array[$i])
            $exactDateTime = "$($key.Split('T')[0].Replace('-','/')) $($key.Split('T')[1].Split('.')[0])"
            $dates[[DateTime]::ParseExact($exactDateTime, "yyyy/MM/dd HH:mm:ss", $Null)] = $array[$eventID].Split(':')[1].Trim()
        }
    }
}

$path = Join-Path (Convert-Path .) "ServerBackupDuration.csv"
Add-Content -Path $path -Value '"HostName","Average","Maximum","Minimum"' -Encoding UTF8

foreach($server in $servers) {
    if(!(Test-Connection $server.Name -Count 1 -Quiet)) {
        Write-Host "$($server.Name) に接続できません"
        continue 
    }

    $taskName = ""
    $isWSB = $True

    if($server.Name -eq "serverName") {
        $taskName = '\Backup'
        $starts = GetEvents -taskName $taskName -serverName $server.Name -eventID '100'
        $ends = GetEvents -taskName $taskName -serverName $server.Name -eventID '102'
        $isWSB = $False
    } else {
        $starts = GetWindowsServerBackupEvents -serverName $server.Name -eventID '1'
        $ends = GetWindowsServerBackupEvents -serverName $server.Name -eventID '4'
    }

    $dates = @{}

    if($Null -eq $starts -or $Null -eq $ends) {
        Write-Host "$($server.Name) is Null"
        continue;
    }

    PushEvent -array $starts
    PushEvent -array $ends

    $dates = $dates.GetEnumerator() | Sort-Object -Property Key

    $diffs = @()

    $isFirstTimeNotNull = $false

    if($isWSB) {
        $startID = 1
        $endID = 4
    } else {
        $startID = 100
        $endID = 102
    }
    foreach ($date in $dates.GetEnumerator()) {
        if ($date.Value -eq $startID -and $isFirstTimeNotNull -eq $False) {
            $first = $date.Key
            $isFirstTimeNotNull = $True
            continue
        }

        if ($isFirstTimeNotNull) {
            if ($date.Value -eq $endID) {
                $second = $date.Key
            }
            else {
                $isFirstTimeNotNull = $False
                continue
            }

            $diffs += ($second - $first).Ticks

            $isFirstTimeNotNull = $False
        }
    }

    $stat = $diffs | Measure-Object -Average -Maximum -Minimum
    $ave = [DateTime]([Int64]($stat.Average))
    $max = [DateTime]([Int64]($stat.Maximum))
    $min = [DateTime]([Int64]($stat.Minimum))

    Add-Content -Path $path -Value "$($server.Name),$($ave),$($max),$($min)" -Encoding UTF8
}
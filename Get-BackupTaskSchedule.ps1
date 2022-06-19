$servers = Get-ADComputer -Properties IPv4Address -Filter { OperatingSystem -like "*Server*" } |
Where-Object { $_.Enabled -eq $True } |
Select-Object Name, IPv4Address

[int]$days = 0

function GetEvents([string]$taskName, [string]$serverName, [bool]$isWSB = $True) {
    $eventIDs = @(1,2,3,4,5,8,9,14,17,18,19,20,21,22,23,49,50,52,100,517,518,521,527,528,544,545,546,561,564,612)
    $resultArrays = @()

    $start = (Get-Date).ToUniversalTime().ToString('o')
    $end = (Get-Date).AddDays($days).ToUniversalTime().ToString('o')

    $timeQuery = "*[System[TimeCreated[@SystemTime>='$($end)' and @SystemTime<='$($start)']]]"

    if($isWSB) {
        foreach ($eventID in $eventIDs) {
            $query = "*[System[Provider[@Name='Microsoft-Windows-Backup'] and (EventID=${eventID})]] and " + $timeQuery
            $resultArrays += wevtutil qe Microsoft-Windows-Backup /q:$query /r:$serverName /uni:true /f:text
        }
    }
    else {
        $eventIDs = 100..104
        foreach ($eventID in $eventIDs) {
            $query = "*[EventData[Data[@Name='TaskName']='${taskName}']] and *[System[(EventID=${eventID})]] and " + $timeQuery
            $resultArrays += wevtutil qe Microsoft-Windows-TaskScheduler/Operational /q:$query /r:$serverName /uni:true /f:text
        }
    }

    return $resultArrays
}

function SplitDateString($dateString) {
    [string]$result = $dateString.SubString($dateString.IndexOf(':') + 1, $dateString.Length - $dateString.IndexOf(':') - 1 ).Trim()
    return $result
}

function PushEvent($array) {
    $backupEvent = [PSCustomObject]@{
        Date        = ""
        Computer    = ""
        EventId     = ""
        Description = ""
        LogName     = ""
    }

    $backupEvents = @()

    for($i = 0; $i -lt $array.Length; $i++) {
        $index = $i % 15
        if($index -eq 1) {
            $backupEvent.LogName = SplitDateString($array[$i])
        }
        elseif($index -eq 3) {
            $key = SplitDateString($array[$i])
            $date = "$($key.Split('T')[0].Replace('-','/')) $($key.Split('T')[1].Split('.')[0])"
            $backupEvent.Date = $date
        }
        elseif($index -eq 4) {
            $backupEvent.EventId = SplitDateString($array[$i])
        }
        elseif($index -eq 11) {
            $backupEvent.Computer = SplitDateString($array[$i])
        }
        elseif($index -eq 13) {
            $backupEvent.Description = $array[$i].Trim()
            $eventObj =  [PSCustomObject]@{
                Date = $backupEvent.Date
                Computer    = $backupEvent.Computer
                EventId     = $backupEvent.EventId
                Description = $backupEvent.Description
                LogName     = $backupEvent.LogName
            }
            $backupEvents += $eventObj
            $eventObj = $Null
        }
        else {
        }
    }

    return $backupEvents
}

function GetServerBaskupStatus($serverName, $results) {
    if (!(Test-Connection $serverName -Count 1 -Quiet)) {
        Write-Host "$($serverName) に接続できません"
        continue 
    }

    Write-Host "$($serverName)..."

    $taskName = ""

    if ($serverName -eq "") {
        $taskName = '\Backup'
        $events = GetEvents -taskName $taskName -serverName $serverName -isWSB $False
    }
    else {
        $events = GetEvents -serverName $serverName
    }

    $results += PushEvent -array $events

    return $results
}

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

while (1) {
    $target = Read-Host @"
    $serverNamesString
"@

    if ($target -eq 'q') {
        break
    }

    $daysFlag = $False
    do {
        $days = Read-Host "Days"
        if ($days -gt 0 -and $days -lt 31) {
            $days = - $days
            $daysFlag = $True
        }
    } while (!$daysFlag)

    $results = @()
    if ($target -eq $serverNamesString.Count - 2) {
        foreach ($server in $servers) {
            $results = GetServerBaskupStatus $server.Name $results
        }
    }
    else {
        $targetName = $serverNames[$target]
        $results = GetServerBaskupStatus $targetName $results
    }

    $results | Sort-Object -Property Date -Descending | Out-GridView
}
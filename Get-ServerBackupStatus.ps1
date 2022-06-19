$servers = Get-ADComputer -Filter { OperatingSystem -like "*Server*" } |
Where-Object { $_.Enabled -eq $True -and $_.Name -notlike "*MCAD*"} |
Select-Object Name

$OutputEncoding = [Console]::OutputEncoding

$now = Get-Date -Format "yyyyMMddHHmmss"
$fileName = "ServerBackupStatus_$now.csv"
$targetFile = Join-Path $Env:TEMP $fileName

Add-Content -Path $targetFile -Value '"HostName","NextRunTime","Status","LastRunTime","LastResult","ScheduledTaskState","RunAsUser","ScheduleType","StartTime","StartDate","Days","Months","","","",""'

function SplitWSBStatusString([string]$statusString) {
    $statusStringValues = $statusString.Split(':')
    $resultString = ""
    for ($i = 1; $i -lt $statusStringValues.Length; $i++) {
        $resultString += $statusStringValues[$i] + ":"
    }
    $resultString = $resultString.Substring(0, $resultString.LastIndexOf(':'))

    return $resultString.Trim()
}

function WriteBackupTask($backupTask) {
    $HostName           = SplitWSBStatusString($backupTask[2]);
    $NextRunTime        = SplitWSBStatusString($backupTask[4]);
    $Status             = SplitWSBStatusString($backupTask[5]);
    $LastRunTime        = SplitWSBStatusString($backupTask[7]);
    $LastResult         = SplitWSBStatusString($backupTask[8]);
    $ScheduledTaskState = SplitWSBStatusString($backupTask[13]);
    $RunAsUser          = SplitWSBStatusString($backupTask[16]);
    $ScheduleType       = SplitWSBStatusString($backupTask[20]);
    $StartTime          = SplitWSBStatusString($backupTask[21]);
    $StartDate          = SplitWSBStatusString($backupTask[22]);
    $Days               = SplitWSBStatusString($backupTask[24]);
    $Months             = SplitWSBStatusString($backupTask[25]);

    Add-Content -Path $targetFile -Value "$($HostName),$($NextRunTime),$($Status),$($LastRunTime),$($LastResult),$($ScheduledTaskState),$($RunAsUser),$($ScheduleType),$($StartTime),$($StartDate),$($Days),$($Months)"
}

foreach($server in $servers) {
    
    $connection = Test-Connection $server.Name -Quiet -Count 1
    if($connection -eq $False) {
        Write-Host "$($server.Name) 接続できませんでした"
        continue;
    }

    $backupTask = schtasks /s $server.Name /query /tn \Microsoft\Windows\Backup\Microsoft-Windows-WindowsBackup /v /fo list

    if($backupTask) {
        WriteBackupTask -backupTask $backupTask
    } else {
        $batTask = schtasks /s $server.Name /query /tn "backup" /v /fo list
        WriteBackupTask -backupTask $batTask
    }
}

Import-Csv -Path $targetFile | Out-GridView -Title "各サーバーのバックアップ状況"

$answer = Read-Host @"
$($targetFile) を生成しました。

この .csv を削除する : y
これまで作成したすべての .csv を削除する : a
この .csv を削除しない : n
"@

if($answer -eq 'a') {
    Remove-Item (Join-Path $Env:TEMP "ServerBackupStatus*")
} elseif($answer -eq 'y') {
    Remove-Item $targetFile
}
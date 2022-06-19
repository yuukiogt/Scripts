$CurrentDir = Split-Path $MyInvocation.MyCommand.Path -Parent
$ScriptName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
$LogFile = Join-Path $CurrentDir ($ScriptName + ".log")

function Log($message) {
    $now = Get-Date
    $log = $now.ToString("yyyy/MM/dd HH:mm:ss.fff") + "`t"
    $log += $message

    Write-Output $log | Out-File -FilePath $LogFile -Encoding UTF8 -append

    return $log
}

try {

    $Path = ".\data.accdb"
    $Table = "table"
    $mutex = $Null

    $connection = New-Object -ComObject ADODB.Connection

    if(-not (Test-Path $Path)) { throw 1 }

    if (([System.IO.Path]::GetExtension($Path)) -eq ".mdb") {
        $connectionString = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = $Path"
    }
    elseif (([System.IO.Path]::GetExtension($Path)) -eq ".accdb") {
        $connectionString = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = $Path"
    }
    else {
        throw 1
    }
    
    $connection.Open($connectionString)
    $query = "SELECT * FROM $Table;"
    $recordSet = New-Object -ComObject ADODB.Recordset

    $adOpenStatic = 3
    $adLockOptimistic = 3
    $recordSet.Open($query, $connection, $adOpenStatic, $adLockOptimistic)

    $recordSet.MoveLast()
    $historyCount = 21

    while ($True) {
        if ($recordSet.Fields.Item("workerCode").Value -eq 642) {
            $workDate = ($recordSet.Fields.Item("date").Value).ToString("yyyy-MM-dd (ddd)")
            $workCode = $recordSet.Fields.Item("workCode").Value
            $workTime = $recordSet.Fields.Item("workTime").Value
            Write-Host "$($workDate) $($workCode) $($workTime)"

            $historyCount--
        }

        if ($True -eq $recordSet.BOF) {
            break;
        }

        if($historyCount -eq 0 ) {
            break;
        }

        $recordSet.MovePrevious()
    }

    $Date = (Get-Date).AddDays(-1).ToString("yyyy-MM-dd (ddd)")
    $tmpDate = Read-Host "Date (default:$Date) "
    if ($tmpDate -eq [string]::Empty) {
        $Date = (Get-Date).AddDays(-1).ToShortDateString()
    } else {
        $Date = ([DateTime]$tmpDate).ToShortDateString()
    }

    $ZZ02 = 480
    $ZZ02 = Read-Host "Task (default $ZZ02) "
    if ($ZZ02 -eq [string]::Empty) {
        $ZZ02 = 480
    }

    $Z100 = 20
    $Z100 = Read-Host "Task (default $Z100) "
    if ($Z100 -eq [string]::Empty) {
        $Z100 = 20
    }

    $C010 = 0
    $C010 = Read-Host "Task (default $C010) "
    if ($C010 -eq [string]::Empty) {
        $C010 = 0
    }

    $B030 = 0
    $B030 = Read-Host "Task (default $B030) "
    if ($B030 -eq [string]::Empty) {
        $B030 = 0
    }

    $now = (Get-Date)

    $mutexName = "Mutex_SetDailyReport"
    $mutex = New-Object System.Threading.Mutex($False, $mutexName)
    $mutex.WaitOne()

    if($ZZ02 -ne 0) {
        $recordSet.AddNew()
        $recordSet.Fields.Item("date") = $Date
        $recordSet.Fields.Item("workerCode") = "642"
        $recordSet.Fields.Item("departCode") = "1140"
        $recordSet.Fields.Item("workCode") = "ZZ02"
        $recordSet.Fields.Item("workTime") = $ZZ02
        $recordSet.Fields.Item("created") = $now
        $recordSet.Fields.Item("modified") = $now
        $recordSet.Update()
    }

    $mutex.ReleaseMutex()
    $mutex = $Null

} catch {
    Log($_.Exception.Message)
    pause
} finally {
    if ($Null -ne $mutex) {
        $mutex.ReleaseMutex()
    }

    if ($Null -ne $recordSet) {
        $recordSet.Close()
    }

    if ($Null -ne $connection) {
        $connection.Close()
    }

    $recordSet = $Null
    $connection = $Null
}
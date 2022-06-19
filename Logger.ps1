
[script]$LogFilePath = $null

function SetLogFilePath($path) {
    [script]$LogFilePath = $path
}

function WriteLog($message) {
    $now = Get-Date
    $log = $now.ToString("yyyy/MM/dd HH:mm:ss.fff") + "`t"
    $log += $message

    Write-Output $log | Out-File -FilePath [script]$LogFilePath -Encoding UTF8 -append

    return $log
}

function DeleteLog() {
    if (Test-Path [script]$LogFilePath) {
        Remove-Item [script]$LogFilePath -Force
    }
}
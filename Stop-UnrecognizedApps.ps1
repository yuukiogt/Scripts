
$processName = "processName"

$process = Get-Process -Name $processName -ErrorAction SilentlyContinue

if ($process) {
    Get-Process -Name $processName | ForEach-Object { $_.ProcessorAffinity = 1; $_.PriorityClass = "Idle" }
    Stop-Process $process.Id -Force
}
$daysAgo = (Get-Date).AddDays(-30)

$ReliabilityStabilityMetrics = Get-CimInstance -ClassName win32_reliabilitystabilitymetrics -filter "TimeGenerated -gt $($daysAgo)" |
Select-Object PSComputerName, SystemStabilityIndex, TimeGenerated

$ReliabilityRecords = Get-CimInstance -ClassName win32_reliabilityRecords -filter "TimeGenerated -gt $($daysAgo)" |
Select-Object PSComputerName, EventIdentifier, LogFile, Message, ProductName, RecordNumber, SourceName, TimeGenerated

$ReliabilityStabilityMetrics |
Export-CSV $env:USERPROFILE\Documents\ReliabilityStabilityMetrics.csv -Encoding UTF8 -NoTypeInformation

$ReliabilityRecords |
Export-CSV $env:USERPROFILE\Documents\ReliabilityRecords.csv -Encoding UTF8 -NoTypeInformation
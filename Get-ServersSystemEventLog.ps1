$ExportFolder = ".\EventLog"

Get-EventLog System -After (Get-Date).AddDays(-1) |
Select-Object EntryType, TimeGenerated, Source, EventID, Category, Message |
Export-CSV -Encoding UTF8 "$ExportFolder\Sys-EventLog.csv"
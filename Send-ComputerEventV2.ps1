$ConfirmPreference = "None"
$ErrorActionPreference = "SilentlyContinue"
$ErrorActionPreference = "Inquire"
$DebugPreference = "SilentlyContinue"
$InformationPreference = "SilentlyContinue"
$ProgressPreference = "SilentlyContinue"
$WarningPreference = "SilentlyContinue"
$VerbosePreference = "SilentlyContinue"

$HostName = [Net.Dns]::GetHostName()
$UserName = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

if ($UserName.Contains("\")) {
    $UserName = $UserName.Split("\")[1]
}

Unregister-ScheduledTask -TaskName Send-ComputerEvent -AsJob

$tempPath = ".\tempPath"
if (!(Test-Path $tempPath)) {
    exit
}

$now = Get-Date -Format "yyyyMMddHHmmss"
$FileNameA = $HostName + "_" + $UserName + "_Events_v2_$now.csv"
$TargetFileE = Join-Path $tempPath $FileNameA

$prevFileLikeQuote = "*$($UserName)_Events_*.csv.txt"

$GetPrevFile = {
    Get-ChildItem -LiteralPath $tempPath |
    Where-Object { $_ -like $prevFileLikeQuote } |
    Sort-Object -Property CreationTime -Descending |
    Select-Object -First 1
}

$prevFile = & $GetPrevFile

$prevDateTime = (Get-Date).AddDays(-30)
if ($Null -ne $prevFile) {
    $prevDateTime = $prevFile.CreationTime
}

$Events = Get-EventLog -LogName System -After $prevDateTime
Where-Object {
    $_.eventID -eq 6005 -or
    $_.eventID -eq 6006 -or
    $_.eventID -eq 6008 -or
    $_.eventID -eq 7001 -or
    $_.eventID -eq 7002 -or
    $_.eventID -eq 20267 -or
    $_.eventID -eq 20268
} |
Select-Object TimeGenerated, eventID

if($Events.Count -gt 0) {
    Add-Content -Path $TargetFileE -Value '"HostName","UserName","DateTime","EventId","Description"' -Encoding UTF8
}

Foreach ($event in $Events) {
    $Description = [string]::Empty
    
    switch ($event.eventID) {
        6005 { $Description = "スタートアップ" }
        6006 { $Description = "シャットダウン" }
        6008 { $Description = "異常終了" }
        7001 { $Description = "サインイン" }
        7002 { $Description = "サインアウト" }
        20267 { $Description = "VPN 接続" }
        20268 { $Description = "VPN 切断" }
    }

    if($Description -ne [string]::Empty) {
        $DateTime = ($event.TimeGenerated).ToString("yyyy-MM-dd HH:mm:ss")
        Add-Content -Path $TargetFileE -Value "$($HostName),$($UserName),$($DateTime),$($event.eventID),$($Description)"  -Encoding UTF8
    }
}

$clientID = ""
$tenantName = "tenant.onmicrosoft.com"
$clientSecret = ""

$reqTokenBody = @{
    grant_type = "client_credentials"
    scope = "https://graph.microsoft.com/.default"
    client_id = $clientID
    client_secret = $clientSecret
}

$tokenRes = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantName/oauth2/v2.0/token" -Method Post -Body $reqTokenBody
$headerParams = @{Authorization = "Bearer $($tokenRes.access_token)" }

$eventFiles = {
    Get-ChildItem -LiteralPath $tempPath |
    Where-Object { $_ -like "*_Events_*.csv" } |
    Sort-Object -Property CreationTime
}

$eventFiles = & $eventFiles
$eventFilesCount = $eventFiles.Count

$driveID = ""

foreach ($file in $eventFiles) {
    $url = "https://graph.microsoft.com/v1.0/sites/tenant.sharepoint.com/drives/$($driveID)/root:/$($file.Name):/content"
    try {
        Invoke-RestMethod -Uri $url -Headers $headerParams -Method Put -InFile $file.FullName -ContentType 'multipart/form-data'
    
        if($eventFilesCount -ne 1) {
            Remove-Item $file.FullName -Force
            $eventFilesCount--
        }
        else {
            Remove-Item "$($tempPath)\*$($UserName)_Events_*.csv.txt"
            Rename-Item $file.FullName -NewName "$($file.FullName).txt"
        }
    }
    catch {
    }
}
$clientID = ""
$tenantName = "tenant.onmicrosoft.com"
$clientSecret = ""

$reqTokenBody = @{
    grant_type    = "client_credentials"
    scope         = "https://graph.microsoft.com/.default"
    client_id     = $clientID
    client_secret = $clientSecret
}

$tokenRes = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantName/oauth2/v2.0/token" -Method Post -Body $reqTokenBody
$headerParams = @{Authorization = "Bearer $($tokenRes.access_token)" }

$srcBasePath = "https://graph.microsoft.com/v1.0/sites/tenant.sharepoint.com:/sites/site"
$dstBasePath = ".\SharePoint"

$siteRes = Invoke-RestMethod -uri $srcBasePath -Method Get -Headers $headerParams
$siteID = $siteRes.id.Split(",")[1]
$drivesRes = Invoke-RestMethod -uri "https://graph.microsoft.com/v1.0/sites/$($siteID)/drives" -Method Get -Headers $headerParams

$sitePath = Join-Path $backupPath $site.Title
New-Item $sitePath -ItemType Directory -Force

foreach ($drive in $drivesRes.value) {
    $driveUrl = "https://graph.microsoft.com/v1.0/drives/$($drive.id)/root/children"
    $childrenRes = Invoke-RestMethod -uri $driveUrl -Method Get -Headers $headerParams

    if ($childrenRes.Value.Length -eq 0) {
        continue
    }

    $drivePath = Join-Path $sitePath $drive.name
    New-Item $drivePath -ItemType Directory -Force

    foreach($child in $childrenRes.value) {
        Write-Host $child.name $child.size "byte"
        Measure-Command {
            $fileUrl = "https://graph.microsoft.com/v1.0/drives/$($drive.id)/items/$($child.id)?select=@microsoft.graph.downloadUrl"
            [string]$fileRes = Invoke-RestMethod -uri $fileUrl -Method Get -Headers $headerParams
            
            $startIndex = $fileRes.IndexOf("downloadUrl=") + ("downloadUrl=").Length
            $endIndex = $fileRes.Length - $startIndex - 1
            $downloadUrl = $fileRes.Substring($startIndex, $endIndex)

            $ProgressPreference = 'SilentlyContinue'
            $downloadRes = Invoke-WebRequest -UseBasicParsing $downloadUrl -OutFile (Join-Path ".\SharePoint\site\temp" $child.name)
        }
    }
}

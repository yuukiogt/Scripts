$ErrorActionPreference = "Inquire"

$credential = Get-Credential
Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
Connect-SPOService -Url https://tenant-admin.sharepoint.com -Credential $credential | Out-Null

$sites = Get-SPOSite -Limit All |
Select-Object LastContentModifiedDate, Status, Title, Url |
Where-Object { $_.Status -eq "Active" -and $_.Title -ne "" }

[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime") | Out-Null

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

function InvokeRestRequest($uri) {
    try {
        $res = Invoke-RestMethod -uri $uri -Method Get -Headers $headerParams
        return $res
    }
    catch {
        $_.Exception.Message
    }
}

$VersionsToKeep = 10

ForEach ($site in $sites) {
    $siteName = $site.Url.Split("/")[4]

    Write-Host -f Yellow $siteName

    $siteUrl = "https://graph.microsoft.com/v1.0/sites/tenant.sharepoint.com:/sites/" + $siteName
    $siteRes = InvokeRestRequest -uri $siteUrl
    $siteID = $siteRes.id.Split(",")[1]

    $listsRes = InvokeRestRequest -uri "https://graph.microsoft.com/v1.0/sites/$($siteID)/lists"
    
    foreach ($list in $listsRes.value) {
        $itemsRes = InvokeRestRequest -uri "https://graph.microsoft.com/v1.0/sites/$($siteID)/lists/$($list.id)/items"
        foreach($item in $itemsRes.value) {
            $versionsRes = InvokeRestRequest -uri "https://graph.microsoft.com/v1.0/sites/$($siteID)/lists/$($list.id)/items/$($item.id)/versions"
            $versions = $versionsRes.value
            if ($versions.Length -gt $VersionsToKeep) {
                $versions = $versions | Sort-Object @{e = { $_.id -as [int] } } -Descending | Select-Object id -Last ($versions.Length - $VersionsToKeep)

                foreach($version in $versions) {
                    try {
                        $res = Invoke-RestMethod -uri "https://graph.microsoft.com/v1.0/sites/$($siteID)/lists/$($list.id)/items/$($item.id)/versions/$($version.id)" -Method Delete -Headers $headerParams
                    }
                    catch {
                        $_.Exception.Message
                    }
                }
            }
        }
    }
}



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

$currentDir = Split-Path $MyInvocation.MyCommand.Path -Parent
$passwordFile = Join-Path $CurrentDir "administrator.txt"
$userName = "administrator@tenant.onmicrosoft.com"
$securePassword = Get-Content $passwordFile | ConvertTo-SecureString -Key (1..16)
$credential = New-Object System.Management.Automation.PSCredential $userName, $securePassword
Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
Connect-SPOService -Url https://tenant-admin.sharepoint.com -Credential $credential

$sites = Get-SPOSite -Limit All | Select-Object LastContentModifiedDate, Status, Title, Url

Disconnect-SPOService

$backupPath = ".\SharePoint"

$siteNames = @()
foreach ($site in $sites) {
    if ($site.Status -ne "Active") {
        continue
    }

    $siteName = $site.Title
    if ([string]::IsNullOrEmpty($siteName)) {
        continue
    }

    $siteNames += $siteName
}

$siteFolders = Get-ChildItem -Path $backupPath -Directory
foreach ($siteFolder in $siteFolders) {
    if (!$siteNames.Contains($siteFolder.Name)) {
        Remove-Item $siteFolder.FullName -Recurse -Force
    }
}

function DownloadFiles($targetPath) {
    $files = Get-PnPFolderItem -FolderSiteRelativeUrl $targetPath -ItemType File
    foreach ($file in $files) {
        $srcPath = Join-Path $targetPath $file.Name
        $dstPath = Join-Path $backupPath $srcPath
        Get-PnPFile -Url $srcPath -Path $dstPath -FileName $file.Name -AsFile
    }

    $folders = Get-PnPFolderItem -FolderSiteRelativeUrl $targetPath -ItemType Folder
    foreach($folder in $folders) {
        $folderPath = Join-Path $targetPath $folder.Name
        New-Item $folderPath -ItemType Directory -Force
        DownloadFiles $folderPath
    }
}

foreach ($site in $sites) {
    if ($site.Status -ne "Active") {
        continue
    }

    $siteName = $site.Url.Split("/")[4]

    if ([string]::IsNullOrEmpty($siteName)) {
        continue
    }

    Connect-PnPOnline -Url "https://tenant.sharepoint.com/sites/site" -Credentials $credential # $site.Url -Credentials $credential
    if($Null -eq $connection) {
        continue
    }

    $documentLibraries = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 -and $_.Hidden -eq $False }

    if($documentLibraries.Length -eq 0) {
        continue
    }

    $sitePath = Join-Path $backupPath $site.Title
    New-Item $sitePath -ItemType Directory -Force

    foreach($documentLibrary in $documentLibraries) {
        $documentLibraryPath = Join-Path $sitePath $documentLibrary.Title
        New-Item $documentLibraryPath -ItemType Directory -Force
        DownloadFiles $documentLibrary.Title
    }

    Disconnect-PnPOnline -Connection $connection
}
$ErrorActionPreference = "SilentlyContinue"
$ProgressPreference = "SilentlyContinue"

$currentDir = Split-Path $MyInvocation.MyCommand.Path -Parent
$scriptName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
$LogFile = Join-Path $currentDir ($scriptName + ".log")

function Log($message) {
    $now = Get-Date
    $log = $now.ToString("yyyy/MM/dd HH:mm:ss.fff") + "`t"
    $log += $message

    Write-Output $log | Out-File -FilePath $LogFile -Encoding UTF8 -append

    return $log
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

function InvokeRestRequest($uri) {
    try {
        $res = Invoke-RestMethod -uri $uri -Method Get -Headers $headerParams
        return $res
    } catch [Microsoft.PowerShell.Commands.HttpResponseException] {
        $statusCode = $_.Exception.Response.StatusCode

        if ($statusCode -eq 429) {
            [int] $delay = [int](($_.Exception.Response.Headers | Where-Object Key -eq 'x-ms-retry-after-ms').Value[0])
            Log("Retry Caught, delaying $delay ms")
            Start-Sleep -Milliseconds $delay
            InvokeRestRequest($uri)
        }
        else {
            Log($_.Exception.Message)
            Log("StatusCode: $($statusCode))")
        }
    } catch {
        Log($_.Exception.Message)
    }
}

$trashBoxPath = ".\SharePoint_TrashBox"

function RemoveDeletedItem($srcItems, $dstPath, $selectParameter = "Name") {
    if($Null -eq $srcItems) {
        return
    }
    $srcItems = $srcItems | Select-Object $selectParameter |
    Out-String -Stream | Select-Object -Skip 3 | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }

    if($selectParameter -eq "Url") {
        $srcItems = $srcItems | ForEach-Object { $_.Split("/")[4] }
    }

    $dstItems = Get-ChildItem -Path $dstPath | Select-Object Name |
    Out-String -Stream | Select-Object -Skip 3 | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }

    $dstItems | Where-Object { $srcItems -notcontains $_ } |
    ForEach-Object {
        $target = Join-Path $dstPath $_
        Set-ItemProperty -Path $target -Name LastWriteTime -Value (Get-Date)
        Move-Item $target -Destination $trashBoxPath -Force
        Log("Moved: $target")
    }
}

$documentLibraryFolders = @{}

$dstBasePath = ".\SharePoint"
$driveID = [string]::Empty

function DownloadFiles($srcChildItems, $dstTargetPath) {
    RemoveDeletedItem -srcItems $srcChildItems -dstPath $dstTargetPath

    foreach($srcChildItem in $srcChildItems) {
        $srcChildItemName = $srcChildItem.name
        if ($srcChildItem.name.Contains("/")) {
            $srcChildItemName = $srcChildItem.name.Replace("/", "-")
        }

        $dstItemPath = Join-Path $dstTargetPath $srcChildItemName

        $isContainBracket = $False
        if($dstItemPath.Contains("[") -or $dstItemPath.Contains("]")) {
            $isContainBracket = $True
            $dstItemPath = [WildcardPattern]::Escape($dstItemPath)
        }

        if ($Null -eq $srcChildItem.folder) {
            if(Test-Path $dstItemPath) {
                if ((Get-ItemProperty $dstItemPath).LastWriteTime -eq $srcChildItem.lastModifiedDateTime) {
                    if ((Get-Item $dstItemPath).Length -eq $srcChildItem.size -or (Get-Item $dstItemPath).Length - $srcChildItem.size -le 8 -or (Get-Item $dstItemPath).Length - $srcChildItem.size -gt -8) {
                        continue
                    }
                }
            }

            $fileUrl = "https://graph.microsoft.com/v1.0/drives/$($driveID)/items/$($srcChildItem.id)?select=@microsoft.graph.downloadUrl"
            [string]$fileRes = InvokeRestRequest -uri $fileUrl
            
            $downloadURLIndex = $fileRes.IndexOf("downloadUrl=")
            if ($downloadURLIndex -eq -1) {
                Log("no downloadURL: $dstItemPath")
                continue
            }

            $startIndex = $downloadURLIndex + ("downloadUrl=").Length
            $endIndex = $fileRes.Length - $startIndex - 1
            $downloadUrl = $fileRes.Substring($startIndex, $endIndex)

            try {
                $downloadRes = Invoke-WebRequest $downloadUrl -OutFile $dstItemPath | Out-Null
            }
            catch [Microsoft.PowerShell.Commands.HttpResponseException] {
                $statusCode = $_.Exception.Response.StatusCode

                if ($statusCode -eq 429) {
                    [int] $delay = [int](($_.Exception.Response.Headers | Where-Object Key -eq 'x-ms-retry-after-ms').Value[0])
                    Log("Retry Caught, delaying $delay ms")
                    Start-Sleep -Milliseconds $delay
                    Invoke-WebRequest $downloadUrl -OutFile $dstItemPath | Out-Null
                }
                else {
                    Log($_.Exception.Message)
                    Log("StatusCode: $($statusCode))")
                }
            }
            catch {
                Log($_.Exception.Message)
                continue
            }

            if($isContainBracket) {
                Rename-Item -LiteralPath $dstItemPath ($dstItemPath -replace '`')
            }

            Set-ItemProperty -Path $dstItemPath -Name LastWriteTime -Value $srcChildItem.lastModifiedDateTime
        }
        else {
            New-Item $dstItemPath -ItemType Directory -Force | Out-Null
            $documentLibraryFolders[$dstItemPath] = $srcChildItem.id
        }
    }

    foreach ($folder in $documentLibraryFolders.GetEnumerator()) {
        $folderUrl = "https://graph.microsoft.com/v1.0/drives/$($driveID)/items/$($folder.Value)/children"
        $itemsRes = InvokeRestRequest -uri $folderUrl
        $documentLibraryFolders.Remove($folder.Key)
        DownloadFiles $itemsRes.value $folder.Key
    }
}

$adminFile = Join-Path $currentDir "administrator.txt"
$userName = "administrator@tenant.onmicrosoft.com"
$securePassword = Get-Content $adminFile | ConvertTo-SecureString -Key (1..16)
$credential = New-Object System.Management.Automation.PSCredential $userName, $securePassword
Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
Connect-SPOService -Url https://tenant-admin.sharepoint.com -Credential $credential | Out-Null

$sites = Get-SPOSite -Limit All |
Select-Object LastContentModifiedDate, Status, Title, Url |
Where-Object { $_.Status -eq "Active" -and $_.Title -ne "" }

RemoveDeletedItem -srcItems $sites -dstPath $dstBasePath -selectParameter "Url"

$siteIDs = @{}

foreach($site in $sites) {
    $siteName = $site.Url.Split("/")[4]

    $documentLibraryFolders = @{}

    $siteUrl = "https://graph.microsoft.com/v1.0/sites/tenant.sharepoint.com:/sites/" + $siteName
    $siteRes = InvokeRestRequest -uri $siteUrl
    $siteID = $siteRes.id.Split(",")[1]

    $siteTitle = $site.Title
    if ($site.Title.Contains("/")) {
        $siteTitle = $site.Title.Replace("/", "-")
    }

    if($siteIDs.ContainsValue($siteID)) {
        Log("$siteTitle ($siteName) : ContainsValue($siteID)")
        continue
    }
    else {
        $siteIDs[$siteTitle] = $siteID
    }

    $drivesRes = InvokeRestRequest -uri "https://graph.microsoft.com/v1.0/sites/$($siteID)/drives"

    $dstSitePath = Join-Path $dstBasePath $siteName

    New-Item $dstSitePath -ItemType Directory -Force | Out-Null

    RemoveDeletedItem -srcItems $drivesRes.Value -dstPath $dstSitePath

    foreach($drive in $drivesRes.value) {
        $driveID = $drive.id

        $driveUrl = "https://graph.microsoft.com/v1.0/drives/$($drive.id)/root/children"
        $driveChildrenRes = InvokeRestRequest -uri $driveUrl

        $driveName = $drive.name
        if ($drive.name.Contains("/")) {
            $driveName = $drive.name.Replace("/", "-")
        }

        $dstDrivePath = Join-Path $dstSitePath $driveName
        New-Item $dstDrivePath -ItemType Directory -Force | Out-Null

        Log("Download... $dstDrivePath")
        DownloadFiles $driveChildrenRes.value $dstDrivePath
    }
}

$dstFolders = @{}
Get-ChildItem $dstBasePath -Recurse -Directory | Select-Object FullName |
ForEach-Object {
    $dstFolders[$_.FullName] = ($_.FullName.Split("\")).Length
}

$dstFolders = $dstFolders.GetEnumerator() | Sort-Object -Property Value -Descending

foreach ($folder in $dstFolders.GetEnumerator()) {
    $childItems = Get-ChildItem $folder.Key | Where-Object {$_.Length -ne 0} | Sort-Object LastWriteTime -Descending
    $latestItemWriteTime = $childItems[0].LastWriteTime
    Set-ItemProperty -Path $folder.Key -Name Attributes -Value "Normal"
    Set-ItemProperty -Path $folder.Key -Name LastWriteTime -Value $latestItemWriteTime
}

$rootChildItems = Get-ChildItem $dstBasePath | Sort-Object LastWriteTime -Descending
$latestItemWriteTime = $rootChildItems[0].LastWriteTime
Set-ItemProperty -Path $dstBasePath -Name LastWriteTime -Value $latestItemWriteTime

Get-ChildItem $trashBoxPath |
ForEach-Object {
    $lastWriteTime = Get-ItemProperty -Path $_.FullName -Name LastWriteTime
    if($lastWriteTime.LastWriteTime -lt (Get-Date).AddDays(-30)) {
        Remove-Item -Path $_.FullName -Recurse -Force
        Log("Removed: $($_.FullName)")
    }
}
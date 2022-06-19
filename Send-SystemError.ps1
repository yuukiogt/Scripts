$folder = ".\ServersSystemDiagReport"
$sender = "user@tenant.co.jp"
$targetEMailAddress = "user@tenant.co.jp"

$excludeErrors = @{
}

function GetComputerName($sourceFolder) {
    $xmlPath = Join-Path $sourceFolder.FullName "report.xml"
    if((Test-Path $xmlPath)) {
        $xml = [XML](Get-Content $xmlPath)
        $clientTable = $xml.Report.Section.Table | Where-Object {$_.Name -eq "client"}
        $computerNameData = $clientTable.Item.Data | Where-Object { $_.Name -eq "computer" }
        $computerName = $computerNameData.InnerXml
        return $computerName
    } else {
        return ""
    }
}

$sources = Get-ChildItem $folder -Directory | Sort-Object -Property Name
$targetReports = @{}
ForEach($source in $sources) {
    $words = $source.Name.Split('_')
    if($words.Count -eq 1) {
        $computerName = GetComputerName -sourceFolder $source
        $parent = Split-Path -Path $source.FullName -Parent
        Move-Item $source.FullName -Destination (Join-Path $parent "$($computerName)_$($source.Name)") -Force
    } else {
        $target = $source.Name.Split('_')[0] + '_' + $source.Name.Split('_')[1].Split('-')[0]
        if($targetReports.Contains($target)) {
            Remove-Item (Join-Path $folder $targetReports[$target]) -Recurse -Force -Confirm:$false
        }
        $targetReports[$target] = $source        
    }
}

$sources = Get-ChildItem $folder -Directory | Sort-Object -Property Name
$oldReports = @{}
ForEach ($source in $sources) {
    $target = $source.Name.Split('_')[0]
    if ($oldReports.ContainsKey($target)) {
        $oldReports[$target] += $source
    }
    else {
        $reportArray = @()
        $reportArray += $source
        $oldReports[$target] = $reportArray
    }
}

ForEach ($oldReport in $oldReports.GetEnumerator()) {
    if ($oldReport.Value.Count -gt 7) {
        for($i = 0; $i -lt $oldReport.Value.Count-7; $i++) {
            $target = (Join-Path $folder $oldReport.Value[$i])
            Remove-Item $target -Recurse -Force -Confirm:$false
        }
    }
}

$sourceMap = @{}
ForEach ($source in $sources) {
    $computerName = $source.Name.Split('_')[0]
    $sourceMap[$computerName] = $source
}

$todayString = (Get-Date).ToString("yyyyMMdd")

ForEach ($source in $sourceMap.Values) {
    $xmlPath = Join-Path $source.FullName "report.xml"
    if (!(Test-Path $xmlPath)) {
        continue
    }

    if($source.Name.IndexOf($todayString) -eq -1) {
        continue;
    }

    $computerName = GetComputerName -sourceFolder $source
    $parent = Split-Path -Path $source.FullName -Parent
    $fileName = Join-Path $parent "$($computerName).xml"
    Set-Content -Path $fileName -Encoding UTF8 -Value "" -Force

    $xml = [System.Xml.XmlDocument](Get-Content -Encoding UTF8 -Raw $xmlPath)
    $errorTable = $xml.Report.Section.Table | Where-Object { $_.Name -eq "error" }
    $errorItems = $errorTable.Item
    ForEach($errorItem in $errorItems) {
        $exclude = $False
        ForEach($data in $errorItem.Data) {
            for($i = 0; $i -ne $excludeErrors[$computerName].Count; $i++) {
                if ($data.InnerText.Contains($excludeErrors[$computerName][$i])) {
                    $exclude = $True
                }
            }
        }
        if(!$exclude) {
            ForEach ($data in $errorItem.Data) {
                Add-Content $fileName $data.InnerText -Encoding UTF8
            }
            Add-Content $fileName "`n" -Encoding UTF8
        }
    }

    if([string]::IsNullOrEmpty([string](Get-Content $fileName))) {
        Remove-Item $fileName
    }
}

$targetXMLFile = Join-Path $folder "*.xml"
if((Test-Path $targetXMLFile)) {
    $clientId = ''
    $tenantName = "tenant.onmicrosoft.com"
    $clientSecret = ""

    $reqTokenBody = @{
        grant_type    = "client_credentials"
        scope         = "https://graph.microsoft.com/.default"
        client_id     = $clientID
        client_secret = $clientSecret
    }

    $tokenRes = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantName/oauth2/v2.0/token" -Method Post -Body $reqTokenBody
    Connect-MgGraph -AccessToken $tokenRes.access_token

    $errorFiles = Get-ChildItem $folder | Where-Object {$_.FullName -like "*.xml"}
    
    $body = New-Object System.Text.StringBuilder
    ForEach($errorFile in $errorFiles) {
        $computerName = [System.IO.Path]::GetFileNameWithoutExtension($errorFile.FullName)
        $targetReportFolder = Get-ChildItem $folder -Directory | Sort-Object -Property LastWriteTime -Descending | Where-Object { $_.BaseName -like "$($computerName)*" } | Select-Object -First 1
        $targetReportDocument = Get-ChildItem $targetReportFolder.FullName | Where-Object {$_.FullName -like "*report.html"}
        if ($Null -eq $targetReportDocument) {
            $targetReportDocument = Get-ChildItem $targetReportFolder.FullName | Where-Object { $_.FullName -like "*report.xml" }
        }
        $body.AppendLine('--------<br>')
        $body.AppendLine((Split-Path $errorFile -Leaf))
        $body.AppendLine('<br>')
        $content = Get-Content $errorFile.FullName
        for($i = 0; $i -lt $content.Length; $i++) {
            $body.AppendLine($content[$i])
            $body.AppendLine('<br>')
        }
        $body.AppendLine("<br><`"File://$($targetReportDocument.FullName)`"><br><br>")
    }

    $params = @{
        Message = @{
            Subject = "Server System Errors"
            Body = @{
                ContentType = "HTML"
                Content     = $body.ToString()
            }
            ToRecipients  = @(
                @{ emailAddress = @{ address = $targetEMailAddress } }
            )
        }
        SaveToSentItems = $false
    }

    Send-MgUserMail -UserId $sender -BodyParameter $params

    Disconnect-MgGraph
}
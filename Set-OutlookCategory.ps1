$clientID = ""
$tenantName = "tenant.onmicrosoft.com"
$clientSecret = ""

$reqTokenBody = @{
    grant_type    = "client_credentials"
    scope         = "https://graph.microsoft.com/.default"
    client_id     = $clientID
    client_secret = $clientSecret
}

Connect-AzureAD

$tokenRes = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantName/oauth2/v2.0/token" -Method Post -Body $reqTokenBody

$headerParams = @{
    "Content-Type" = "application/json"
    "Authorization" = "Bearer $($tokenRes.access_token)"
}

$jsonContents = (Get-Content (Join-Path $PSScriptRoot "OutlookCategories.json") -Raw -Encoding UTF8 | ConvertFrom-Json)

$isAll = Read-Host "1:All, 2:User"

if($isAll -eq 1) {
    $csvPath = ".\Users.csv"
    $users = Import-Csv $csvPath | Select-Object DisplayName, MailAddress

    ForEach ($user in $users) {
        $user.DisplayName

        foreach ($jsonContent in $jsonContents.PSObject.Properties) {
            try {
                $body = [System.Text.Encoding]::UTF8.GetBytes(($jsonContent.Value | ConvertTo-Json))
                $res = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$($user.MailAddress)/outlook/masterCategories" -Headers $headerParams -Method Post -Body $body
            }
            catch {
                Write-Host $user.DisplayName
                $res = $_.Exception.Response.GetResponseStream()
                $reader = New-Object System.IO.StreamReader($res)
                $reader.BaseStream.Position = 0
                $reader.DiscardBufferedData()
                $responseBody = $reader.ReadToEnd()
                Write-Host $responseBody
            }
        }
    }
} else {
    $mailAddress = Read-Host "mail address"

    foreach ($jsonContent in $jsonContents.PSObject.Properties) {
        try {
            $body = [System.Text.Encoding]::UTF8.GetBytes(($jsonContent.Value | ConvertTo-Json))
            $res = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$($mailAddress)/outlook/masterCategories" -Headers $headerParams -Method Post -Body $body
        }
        catch {
            Write-Host $user.DisplayName
            $res = $_.Exception.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($res)
            $reader.BaseStream.Position = 0
            $reader.DiscardBufferedData()
            $responseBody = $reader.ReadToEnd()
            Write-Host $responseBody
        }
    }
}

Disconnect-AzureAD
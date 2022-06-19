$ErrorActionPreference = "Inquire"

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

$CsvPath = ".\Users.csv"

$Subject = Read-Host "Subject"
Import-Csv $csvPath | Foreach-Object {
    $UsersRes = InvokeRestRequest -uri "https://graph.microsoft.com/v1.0//users/$($_.MailAddress)/messages?$top=300"
    $TargetMail = $UsersRes.Value | Where-Object {$_.Subject -like "*$($Subject)*"}
    
    if($TargetMail) {
        Write-Host -F Cyan "メールが届いています"
    }
    else {
        Write-Host -F Yellow "メールが届いていません"
    }
}

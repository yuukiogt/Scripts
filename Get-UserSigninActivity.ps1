$ConfirmPreference = "None"
$ErrorActionPreference = "SilentlyContinue"
$DebugPreference = "SilentlyContinue"
$InformationPreference = "SilentlyContinue"
$ProgressPreference = "SilentlyContinue"
$WarningPreference = "SilentlyContinue"
$VerbosePreference = "SilentlyContinue"

$UserName = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

if ($UserName.Contains("\")) {
    $UserName = $UserName.Split("\")[1]
}

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

$driveID = ""

$OU = ""

$users = Get-ADUser -SearchBase "OU=$($OU),DC=tenant,DC=local" -Filter *

foreach($user in $users) {
    $userPrincipalName = "$($user.SamAccountName)@tenant.co.jp"
    $signinUri = "https://graph.microsoft.com/v1.0/auditLogs/signIns"
    try {
        $res = Invoke-RestMethod -Uri $signinUri -Headers $headerParams -Method Get
        $res
    }
    catch {
        $_.Exception.Message
    }
}
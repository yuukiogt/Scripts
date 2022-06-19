
function global:Read-MultiLine ( [string]$prompt, [string]$endChar = ";" ) {
    if ( $prompt.Length -gt 0 ) { Write-Host ($prompt + ":") }
	
    while (1) {
        $ip += Read-Host
        if ($ip.SubString($ip.Length - $endChar.Length) -eq $endChar) { break; }
        else { $ip += "`n" }
    }

    return $ip.Substring(0, $ip.Length - $endChar.Length)
}

try {
    $path = Join-Path (Split-Path $MyInvocation.MyCommand.path) "Send-Inquiry.json"

    if(!(Test-Path $path) -or ((Read-Host "new credential? (y/n)") -eq "y")) {
        $credential = Get-Credential
        ConvertTo-Json @{
            userId   = $credential.UserName;
            password = $credential.Password | ConvertFrom-SecureString;
        } | Set-Content $path
    }

    $json = Get-Content $path | ConvertFrom-Json
    $password = $json.password | ConvertTo-SecureString
    $credential = New-Object System.management.Automation.PsCredential($json.userId, $password)

    $user = $credential.UserName
    $password = $credential.Password
    $from = Read-Host "from"
    $to = Read-Host "to"
    $subject = Read-Host "subject"
    $body = global:Read-MultiLine -prompt "body"

    $mail = New-Object System.Web.Mail.MailMessage
    $mail.From = $from
    $mail.To = $to
    $mail.Subject = $subject
    $mail.Body = $body
    $mail.Fields["http://schemas.microsoft.com/cdo/configuration/sendusing"] = 2
    $mail.Fields["http://schemas.microsoft.com/cdo/configuration/smtpserver"] = "smtp.mail.yahoo.co.jp"
    $mail.Fields["http://schemas.microsoft.com/cdo/configuration/smtpserverport"] = 465
    $mail.Fields["http://schemas.microsoft.com/cdo/configuration/smtpauthenticate"] = 1
    $mail.Fields["http://schemas.microsoft.com/cdo/configuration/sendusername"] = $user
    $mail.Fields["http://schemas.microsoft.com/cdo/configuration/sendpassword"] = $password
    $mail.Fields["http://schemas.microsoft.com/cdo/configuration/smtpusessl"] = $true
    [System.Web.Mail.SmtpMail]::SmtpServer = "smtp.mail.yahoo.co.jp"

    $count = Read-Host "count"
    $wait = Read-Host "wait (s)"

    for($i = 0; $i -lt $count; $i++) {
        [System.Web.Mail.SmtpMail]::Send($mail)
        Start-Sleep -Seconds $wait
    }
} catch {
    $_.Exception.Message
}
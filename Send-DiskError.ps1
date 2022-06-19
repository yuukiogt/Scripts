$datetime = Get-Date

$hostName = "smtp.office365.com"
$port = 587

$from = ""
$to = ""

$subject = "DiskError on $targetHostName"

$smtp = New-Object Net.Mail.SmtpClient
$smtp.Host = $hostName
$smtp.Port = $port
$smtp.EnableSsl = $true

$smtp.Credentials = Get-Credential

$logMessage = Get-EventLog -LogName System -Newest 1 -EntryType Error, Warning | Format-List | Out-String

$body = $datetime.ToString() + "_" + $targetHostName + $logMessage

$smtp.send($from, $to, $subject, $body)
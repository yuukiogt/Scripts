$Credential = Get-Credential

$hostName = "smtp.office365.com"
$port = 587

$from = $userName
$to = $userName

$subject = ""

$smtp = New-Object Net.Mail.SmtpClient
$smtp.Host = $hostName
$smtp.Port = $port
$smtp.EnableSsl = $true

$smtp.Credentials = $Credential
$body = ""
$smtp.send($from, $to, $subject, $body)
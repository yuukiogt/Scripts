$Credential = Get-Credential
Connect-ExchangeOnline -Credential $Credential

$target = ""

Get-ExoMailbox -Identity $target | Out-GridView

Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
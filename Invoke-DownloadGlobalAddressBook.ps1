Connect-ExchangeOnline -Credential Get-Credential

Get-GlobalAddressList | Update-GlobalAddressList -Verbose -WarningAction SilentlyContinue

Disconnect-ExchangeOnline -Confirm:$false
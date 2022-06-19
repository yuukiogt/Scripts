Import-Module ExchangeOnlineManagement

$Credential = Get-Credential
Connect-ExchangeOnline -Credential $Credential

$csvPath = ".\Users.csv"

$Department = "*"

$users = Import-Csv $csvPath |
Where-Object {$_.Department -like $Department -and $_.Enabled -eq 1} |
Select-Object MailAddress

ForEach($user in $users) {
    Add-UnifiedGroupLinks -Identity "salesforce@camnac.onmicrosoft.com" -LinkType Members -Links $user.MailAddress
}

Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
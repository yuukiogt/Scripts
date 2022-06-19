$Credential = Get-Credential
Connect-ExchangeOnline -Credential $Credential

$csvPath = ".\DistributionLists.csv"

Import-Csv $csvPath | Foreach-Object {
    Set-Group -Identity $_.MailAddress -SeniorityIndex $_.SeniorityIndex
    Set-DistributionGroup -Identity $_.MailAddress -ManagedBy "administrator@tenant.onmicrosoft.com"
}

Disconnect-ExchangeOnline -Confirm:$false
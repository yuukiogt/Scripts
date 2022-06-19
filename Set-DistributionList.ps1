$Credential = Get-Credential
Connect-ExchangeOnline -Credential $Credential

$csvPath = ".\DistributionLists.csv"

Import-Csv $csvPath | Foreach-Object {
    Set-Group -Identity $_.MailAddress -DisplayName $_.DisplayName -ManagedBy $_.ManagedBy -SeniorityIndex $_.SeniorityIndex -IsHierarchicalGroup $True
}

Disconnect-ExchangeOnline -Confirm:$false
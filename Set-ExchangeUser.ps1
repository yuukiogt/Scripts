$Credential = Get-Credential
Connect-ExchangeOnline -Credential $Credential

$csvPath = ".\Users.csv"

$target = Read-Host "Identity or all"
if($target -eq 'all') {
    Import-Csv $csvPath | Foreach-Object {
        Set-User -Identity $_.MailAddress -Company $_.Company -Department $_.Department -Manager $_.Manager -Office $_.Office -PhoneticDisplayName $_.PhoneticDisplayName -Title $_.Title -SeniorityIndex $_.SeniorityIndex
    }
} else {
    $user = Import-Csv $csvPath | Where-Object {$_.UserPrincipalName -eq $target}
    Set-User -Identity $user.MailAddress -Company $user.Company -Department $user.Department -Manager $user.Manager -Office $user.Office -PhoneticDisplayName $user.PhoneticDisplayName -Title $user.Title -SeniorityIndex $user.SeniorityIndex
}

Disconnect-ExchangeOnline -Confirm:$false
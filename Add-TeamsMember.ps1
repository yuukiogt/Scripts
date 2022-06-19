Connect-MicrosoftTeams -Credential (Get-Credential)

$csvPath = ".\Users.csv"

$Office = "*"
$users = Import-Csv $csvPath |
Where-Object {$_.Office -like $Office -and $_.Enabled -eq 1} |
Select-Object MailAddress

$GroupId = ""
$DisplayName = ""

ForEach($user in $users) {
    Add-TeamChannelUser -GroupId $GroupId -DisplayName $DisplayName -User $user.MailAddress
}

Disconnect-MicrosoftTeams
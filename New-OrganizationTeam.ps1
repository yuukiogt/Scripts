$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking

Connect-MicrosoftTeams -Credential $UserCredential

$teamsCsvPath = ".\Teams.csv"
$usersCsvPath = ".\Users.csv"
$distributionListCsvPath = ".\DistributionLists.csv"

Import-Csv $teamsCsvPath | Foreach-Object {
    $isExist = Get-Team -DisplayName $_.DisplayName
    
    if($isExist -ne $Null) {
        continue;
    }

    New-Team -DisplayName $_.DisplayName -MailNickName $_.DiplayName -Owner $_.Owner -Visibility "Private" -AllowGuestCreateUpdateChannels $False -AllowGuestDeleteChannels $False
}
    
Import-Csv $usersCsvPath | Foreach-Object {
    if($_.Team1 -ne "") {
        $targetTeam = Get-Team -DisplayName $_.Team1
        Add-TeamUser -GroupId targetTeam.GroupId -User $_.MailAddress -Role $_.TeamRole
    }

    if($_.Team2 -ne "") {
        $targetTeam = Get-Team -DisplayName $_.Team2
        Add-TeamUser -GroupId targetTeam.GroupId -User $_.MailAddress -Role $_.TeamRole
    }
}

Remove-PSSession $Session
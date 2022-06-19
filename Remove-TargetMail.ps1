$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $UserCredential -Authentication "Basic" -AllowRedirection
Import-PSSession $Session -DisableNameChecking

$csvPath = ".\Users.csv"

Import-Csv $csvPath | Foreach-Object {
    Get-Mailbox -RecipientTypeDetails UserMailbox -Identity $_.MailAddress | Search-Mailbox -SearchQuery "subject:'subject', Sent:dd/MM/yyyy" -DeleteContent -Confirm:$false
}

Remove-PSSession $Session
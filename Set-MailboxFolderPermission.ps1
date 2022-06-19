$Credential = Get-Credential

Connect-ExchangeOnline -Credential $Credential

$type = Read-Host "Type 1:User 2:Room 3:Equpment"

switch ($type) {
    1 {
        $target = Read-Host "Identity or all"
        if($target -eq "all") {
            $userMailbox = Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox
            $userMailbox | ForEach-Object {
                $_.DisplayName

                try {
                    Set-MailboxFolderPermission -Identity "$($_.UserPrincipalName):\予定表" -User "default" -AccessRights Reviewer
                }
                catch {
                    Set-MailboxFolderPermission -Identity "$($_.UserPrincipalName):\Calendar" -User "default" -AccessRights Reviewer
                }
            }
        }
        else {
            $userMailbox = Get-EXOMailbox -Identity $target -RecipientTypeDetails UserMailbox
            Set-MailboxFolderPermission -Identity "$($userMailbox.UserPrincipalName):\Calendar" -User "default" -AccessRights Reviewer
        }
    }
    2 {
        $target = Read-Host "Email or all"
        if ($target -eq "all") {
            $roomMailbox = Get-EXOMailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited
            $roomMailbox | ForEach-Object {
                $_.DisplayName
                $calendarPath = (Get-MailboxFolderStatistics $_.Alias | Where-Object { $_.FolderType -eq “Calendar“ }).Identity -replace “\\”, ”:\”
                Set-MailboxFolderPermission -Identity $calendarPath -User “default” -AccessRights Reviewer
            }
        }
        else {
            $roomMailbox = Get-EXOMailbox -Identity -RecipientTypeDetails RoomMailbox
            $roomMailbox.Alias
            $calendarPath = (Get-MailboxFolderStatistics $roomMailbox.Alias | Where-Object { $_.FolderType -eq “Calendar“ }).Identity -replace “\\”, ”:\”
            Set-MailboxFolderPermission -Identity $calendarPath -User “default” -AccessRights Reviewer
        }
    }
    3 {
        $target = Read-Host "Email or all"
        if ($target -eq "all") {
            $equipmentMailbox = Get-Mailbox -RecipientTypeDetails EquipmentMailbox -ResultSize Unlimited
            $equipmentMailbox | ForEach-Object {
                $_.DisplayName
                $calendarPath = $_.Alias + ":\Calendar"
                Set-MailboxFolderPermission -Identity $calendarPath -User “default” -AccessRights Reviewer
            }
        }
        else {
            $equipmentMailbox = Get-Mailbox Identity $target -RecipientTypeDetails EquipmentMailbox
            $equipmentMailbox.Alias
            $calendarPath = $equipmentMailbox.Alias + ":\Calendar"
            Set-MailboxFolderPermission -Identity $calendarPath -User “default” -AccessRights Reviewer
        }
    }
}

Disconnect-ExchangeOnline -Confirm:$false
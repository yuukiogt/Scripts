Connect-ExchangeOnline -Credential Get-Credential

$name = Read-Host "Name"

$mailbox = Get-Mailbox -Name  $name

function GetMailboxProperties() {
    $displayName = Read-Host "DisplayName"
    
    while(1) {
        $type = Read-Host "Type (Room:1, Equipment:2)"
        if($type -eq 1 -or $type -eq 2) {
            break
        }
    }

    $capacity = Read-Host "Capacity"
    $office = Read-Host "Office"

    return @{
        "displayName"=$displayName
        "type"=$type
        "capacity"=$capacity
        "office"=$office
    }
}

if($Null -eq $mailbox) {
    if((Read-Host "$($name)のメールボックスを作成しますか？(y/n)") -eq 'y') {
        $properties = GetMailboxProperties

        if($properties["type"] -eq 1) {
            New-Mailbox -Name $name -Room -DisplayName $properties["displayName"] -ResourceCapacity $properties["capacity"] -Office $properties["office"]
        } else {
            New-Mailbox -Name $name -Equipment -DisplayName $properties["displayName"] -ResourceCapacity $properties["capacity"] -Office $properties["office"]
        }

        Set-CalendarProcessing -Identity $name -AutomateProcessing AutoAccept -RemovePrivateProperty $false
        Set-MailboxFolderPermission "$($name):\予定表" -User '既定' -AccessRights Reviewer        
        Add-MailboxPermission $name -User "ml-its-admin@camnac.co.jp" -AccessRights FullAccess
    }
} else {
    $properties = GetMailboxProperties

    if($properties["type"] -eq 1) {
        Set-Mailbox -Name $name -Room -DisplayName $properties["displayName"] -ResourceCapacity $properties["capacity"] -Office $properties["office"]
    } else {
        Set-Mailbox -Name $name -Equipment -DisplayName $properties["displayName"] -ResourceCapacity $properties["capacity"] -Office $properties["office"]
    }
}

Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
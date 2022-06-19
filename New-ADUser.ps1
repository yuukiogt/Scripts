
$ADUsers = Get-ADUser -Properties EmailAddress -Filter * | Select-Object EmailAddress | Where-Object { $_.EmailAddress -ne $Null } 

[array]$adUsersMailAddresses = $Null

$ADusers | Foreach-Object {
    $adUsersMailAddresses += $_.EmailAddress
}

$usersCsvPath = ".\Users.csv"

[array]$usersMailAddresses = $Null

$csv = Import-Csv $usersCsvPath
$csv | Foreach-Object {
    $usersMailAddresses += $_.MailAddress
}

$newUserMailAddresses = $usersMailAddresses | Where-Object { $adUsersMailAddresses -notcontains $_ }

$newUserMailAddresses

$OU = ""

if ($newUserMailAddresses.Count -ne 0) {
    $answer = Read-Host "これらのユーザーを新規追加しますか？ (y/n)"
    
    if ($answer -eq 'y') {
        foreach ($newUserMailAddress in $newUserMailAddresses) {
            $csv | Foreach-Object {
                if ($newUserMailAddress -eq $_.MailAddress) {
                    $sn = $_.SN
                    $givenName = $_.GivenName
                    $displayName = $_.DisplayName
                    $description = $_.Description
                    $physicalDeliveryOfficeName = $_.Office
                    $mail = $_.MailAddress
                    $sAMAccountName = $_.UserPrincipalName
                    $userPrincipalName = $_.UserPrincipalName + "@domain.local"
                    $info = $_.Info
                    $company = $_.Company
                    $department = $_.Department
                    $title = $_.Title

                    $manager = (Get-ADUser -Identity $_.Manager).DistinguishedName
                    $msDSPhoneticDisplayName = $_.PhoneticDisplayName

                    [array]$ous = $_.Department.Split("/")
                    if($ous.Count -gt 1) {
                        [array]::Reverse($ous)
                    }

                    $ouString = ""

                    foreach($ou in $ous) {
                        $ouString += "OU=" + $ou + ", "
                    }

                    $path = $ouString + "OU=$($OU), DC=domain, DC=local"

                    New-ADUser -Name $DisplayName -Path $path `
                    -Surname $sn `
                    -GivenName $givenName `
                    -DisplayName $displayName `
                    -Description $description `
                    -Office $physicalDeliveryOfficeName `
                    -EmailAddress $mail `
                    -UserPrincipalName $userPrincipalName `
                    -OtherAttributes @{'info' = $info; 'msDS-PhoneticDisplayName' = $msDSPhoneticDisplayName } `
                    -Company $company `
                    -Department $department `
                    -Title $title `
                    -Manager $manager `
                    -SamAccountName $sAMAccountName `
                    -AccountPassword (ConvertTo-SecureString -AsPlainText "initialPassword" -Force) `
                    -ChangePasswordAtLogon $True -Enabled $True -PasswordNeverExpires $False
            
                    for ($i = 0; $i -lt $ous.Count; $i++) {
                        $group = Get-ADGroup $ous[$i]
                        if($group -ne $Null) {
                            break;
                        }
                    }
                    
                    if($group -ne $Null) {
                        Add-ADGroupMember $group $sAMAccountName
                    }
                }
            }   
        }
    }
}
else {
    Write-Host "新規ユーザーは存在しません"
}

Read-Host "Press any key.."
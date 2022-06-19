Connect-MsolService -Credential Get-Credential

$type = Read-Host "1:GetLicenseStatus 2:ShowAllUsersLicense 3:AssignLicense 4:DeleteLicense 5:DeletedUser 6:DeleteUserCompletely"
	
switch($type) {
    1 {
        Get-MsolAccountSku
    }

    2 {
        Get-MsolUser -All | Where-Object { $_.isLicensed -eq $true }
    }

    3 {
        $upn = Read-Host "UserPrincipalName"
        Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses "AccountSkuId"
        Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses "tenant:O365_BUSINESS"
    }

    4 {
        $upn = Read-Host "UserPrincipalName"
        Set-MsolUserLicense -UserPrincipalName $upn -RemoveLicenses "AccountSkuId"
    }

    5 {
        Get-MsolUser -ReturnDeletedUsers -All
    }

    6 {
        $mailAddress = Read-Host "MailAddress"
        Remove-MsolUser -UserPrincipalName $mailAddress -RemoveFromRecycleBin
    }
}
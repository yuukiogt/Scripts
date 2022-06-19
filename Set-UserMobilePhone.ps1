$userCredential = Get-Credential
Connect-ExchangeOnline -Credential $userCredential -ShowProgress $True

$ADUsers = Get-ADUser -Properties Mobile -Filter {company -eq "companyname" -and Enabled -eq $True} |
Select-Object UserPrincipalName, Mobile, SamAccountName

$MSOLUsers = Get-MSOLUser -MaxResults 300 | Select-Object UserPrincipalName, MobilePhone |
Where-Object UserPrincipalName -notlike "eqp*" 

ForEach ($ADUser in $ADUsers) {
    $MSOLUser = $MSOLUsers | Where-Object { $_.UserPrincipalName -eq $ADUser.UserPrincipalName }
    if ([string]::isNullOrEmpty($MSOLUser.UserPrincipalName)) {
        Write-Host "Skip - M365に作成されていません: " $ADUser.SamAccountName
    }
    elseif ($ADUser.mobile -ne $MSOLUser.MobilePhone ) {
        Write-Host "ADの情報でOffice365を上書きします【$ADUser.UserPrincipalName】"
        Write-Host " - ActiveDirectory: ($ADUser.mobile) の値でM365: ($MSOLUser.MobilePhone) の値を上書きします"
        Set-MsolUser -UserPrincipalName $MSOLUser.UserPrincipalName -MobilePhone $ADUser.mobile
        Write-Host " - 完了しました"
    }
}

Read-Host "Press any key.."
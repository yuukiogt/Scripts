Connect-AzureAD -Credential (Get-Credential)

while($True) {
    $targetUser = Read-Host "UserPrincipalName (quit:q)"
    
    if($targetUser -eq 'q') {
        break;
    }

    $foundUsers = Get-AzureADUser -All:$True | Select-Object ObjectId, DisplayName, UserPrincipalName | Where-Object { $_.UserPrincipalName -like "*$($targetUser)*" }

    $foundUsers | ForEach-Object {
        $devices = Get-AzureADUserOwnedDevice -ObjectId $_.ObjectId

        $user = $_

        $devices | ForEach-Object {
            $computer = Get-ADComputer -Identity $_.DisplayName -Properties LastLogonDate
            Write-Host "$($user.DisplayName):`t$($_.DisplayName)`t$($computer.LastLogonDate.ToString("yyyy-MM-dd HH:mm:ss"))"
        }
    }
}

Disconnect-AzureAD
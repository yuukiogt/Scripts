$connection = Connect-AzureAD -Credential (Get-Credential)

$devices= Get-AzureADDevice -All $True | Select-Object DisplayName,ObjectId

$owners = @{}

ForEach($device in $devices) {
    $owner = Get-AzureADDeviceRegisteredOwner -ObjectId $device.ObjectId | Select-Object DisplayName
    $owners[$device.DisplayName] = $owner.DisplayName
}

$owners | Out-GridView

if($connection) {
    Disconnect-AzureAD
}
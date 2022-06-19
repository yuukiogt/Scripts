$friendlyName = Read-Host "Friendly Name "
$device = Get-PnpDevice | Where-Object { $_.friendlyname -like "*$($friendlyName)*" }

$device

if($Null -ne $device) {
    if ((Read-Host "OK? (y/n)") -eq "y") {
        $device | Disable-PnpDevice -Confirm:$True
    }
}
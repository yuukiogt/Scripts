$count = (Get-NetFirewallRule -Name "HNS Container Networking - *").Count

Remove-NetFirewallRule -Name "HNS Container Networking - *"
Write-Host "$count Rules Removed."

if($count -gt 0) {
    $reboot = Read-Host "Reboot (y/n)"

    if($reboot -eq 'y') {
        Restart-Computer -Force
    }
}
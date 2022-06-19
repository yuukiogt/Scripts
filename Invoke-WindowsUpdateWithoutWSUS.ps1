$ErrorActionPreference = "SilentlyContinue"

$IsExistAU = $False
$AUPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"

try {
    $IsExistAU = Test-Path -LiteralPath $AUPath

    if(!$IsExistAU) {
        Exit
    }

    Set-ItemProperty -LiteralPath $AUPath -Name "UseWUServer" -Value 0
    Restart-Service -Name "wuauserv"

    $Updates = Start-WUScan
    if($Updates.Count -gt 0) {
        Install-WUUpdates -Updates $Updates
    }

    if(Get-WUIsPendingReboot) {
        Write-Host "更新を完了するには再起動が必要です。"       
    }
} catch {
    Write-Host $_.Exception.Message
} finally {
    if($IsExistAU) {
        Set-ItemProperty -LiteralPath $AUPath -Name "UseWUServer" -Value 1
    }
}
$ConfirmPreference = "None"
$ErrorActionPreference = "SilentlyContinue"
$DebugPreference = "SilentlyContinue"
$InformationPreference = "SilentlyContinue"
$ProgressPreference = "SilentlyContinue"
$WarningPreference = "SilentlyContinue"
$VerbosePreference = "SilentlyContinue"

try {   
    if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        $argsToAdminProcess = ""
        $Args.ForEach{ $argsToAdminProcess += "`"$PSItem`"" }
        Start-Process powershell.exe "-File `"$PSCommandPath`" $argsToAdminProcess" -Verb RunAs -WindowStyle Hidden
        exit
    }

    Get-Process -Name powershell |
    Where-Object -FilterScript { $_.MainWindowTitle.StartsWith("Windows") } |
    Stop-Process

    while($True) {
        $providerName = Read-Host "DriverProviderName (Default:all, q:quit) "
        if($providerName -eq 'q') {
            break;
        }
        $devices = Get-WmiObject -query "Select * from Win32_PnPSignedDriver" |
        Select-Object DeviceName, DriverProviderName, DriverVersion |
        Where-Object { $_.DriverProviderName -like "*$($providerName)*"}

        $devices | Format-Table DeviceName, DriverProviderName, DriverVersion
    }

} catch {
    $_.Exception.Message
} finally {
    Read-Host "Press any key.."    
}
$ConfirmPreference = "None"
$ErrorActionPreference = "SilentlyContinue"
$DebugPreference = "SilentlyContinue"
$InformationPreference = "SilentlyContinue"
$ProgressPreference = "SilentlyContinue"
$WarningPreference = "SilentlyContinue"
$VerbosePreference = "SilentlyContinue"

$dest = (Join-Path $env:TEMP "Launch-Scan.ps1")
if(!(Test-Path $dest)) {
    Copy-Item "\\dc\SYSVOL\tenand.local\scripts\Launch-Scan.ps1" $dest -Force
}

Start-Process powershell $dest -WindowStyle Hidden
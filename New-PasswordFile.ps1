$Currentdir = Split-Path $MyInvocation.MyCommand.Path -Parent

$SettingsFile = Join-Path $CurrentDir "pass.bin"

$Credential = Get-Credential

$Credential.Password | ConvertFrom-SecureString | Set-Content $SettingsFile
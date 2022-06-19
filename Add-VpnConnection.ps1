$ConfirmPreference = "None"
$ErrorActionPreference = "SilentlyContinue"
$DebugPreference = "SilentlyContinue"
$InformationPreference = "SilentlyContinue"
$ProgressPreference = "SilentlyContinue"
$WarningPreference = "SilentlyContinue"
$VerbosePreference = "SilentlyContinue"

function hasAdminAuth() {
    ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Start-ScriptAsAdmin {
    param([string] $ScriptPath, [object[]] $ArgumentList)
    if (!(hasAdminAuth)) {
        $list = @($ScriptPath)
        if ($null -ne $ArgumentList) {
            $list += @($ArgumentList)
        }
        Start-Process powershell -ArgumentList $list -Verb RunAs -Wait
    }
}

Start-ScriptAsAdmin -ScriptPath $PSCommandPath

$executionPolicy = Get-ExecutionPolicy
if($executionPolicy -ne "RemoteSigned") {
    Set-ExecutionPolicy -Scope Process RemoteSigned -Force
}

$PreKey = ""
$VpnUrl = ""
$VpnName = ""

Add-VpnConnection -Name $VpnName -ServerAddress $VpnUrl -RememberCredential -L2tpPsk $PreKey -AuthenticationMethod Chap,MSChapv2 -EncryptionLevel Required -TunnelType L2tp -Force

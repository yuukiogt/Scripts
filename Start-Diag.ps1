
$ConfirmPreference = "None"
$ErrorActionPreference = "SilentlyContinue"
$DebugPreference = "SilentlyContinue"
$InformationPreference = "SilentlyContinue"
$ProgressPreference = "SilentlyContinue"
$WarningPreference = "SilentlyContinue"
$VerbosePreference = "SilentlyContinue"

if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $argsToAdminProcess = ""
    $Args.ForEach{ $argsToAdminProcess += "`"$PSItem`"" }
    Start-Process powershell.exe "-File `"$PSCommandPath`" $argsToAdminProcess" -Verb RunAs -WindowStyle Hidden
    exit
}

$hostName = [Net.Dns]::GetHostName()
$userName = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

if ($userName.Contains("\")) {
  $userName = $userName.Split("\")[1]
}

$myDocument = [Environment]::GetFolderPath('MyDocuments')
$now = Get-Date -Format "yyyyMMddHHmmss"
$logDir = New-Item -ItemType Directory (Join-Path $myDocument "$($hostName)_$($now)_Diag")
$logFile = Join-Path $logDir "Log.log"
$networkFile= Join-Path $logDir "Network.log"
$firewallFile = Join-Path $logDir "Firewall.log"
$firewallRegistryFile = Join-Path $logDir "FirewallRegistry.log"
$systemInfoFile = Join-Path $logDir "SystemInfo.log"
$msinfo32File = Join-Path $logDir "msinfo32.log"
$userInfoFile = Join-Path $logDir "User.log"
$taskInfoFile = Join-Path $logDir "Task.log"
$appInfoFile = Join-Path $logDir "AppInfo.log"

function WriteLog($message) {
  $now = Get-Date
  $log = $now.ToString("yyyy/MM/dd HH:mm:ss.fff") + "`t"
  $log += $message

  Write-Output $log | Out-File -FilePath $logFile -Encoding UTF8 -Append

  return $log
}

function Write-File {
  [CmdletBinding()]
  param (
      [Parameter(Mandatory=$True, ValueFromPipeline=$True)]
      [ValidateNotNull()]
      $obj,
      [string]$filePath,
      [string]$settingsString
  )

  begin {
    WriteLog("Start $($settingsString)")
  }

  process {
    $obj | Out-File -FilePath $filePath -Encoding UTF8 -Append
  }

  end {
    WriteLog("End $($settingsString)")
  }
}

ipconfig /all | Write-File -filePath $networkFile -settingsString "ipconfig"
Get-NetIPConfiguration -All -Detailed | Write-File -filePath $networkFile -settingsString "Get-NetIPConfiguration"

WriteLog("Start netsh advfirewall export")
$wfwExportPath = (Join-Path $logDir "$($hostName)_firewall.wfw")
netsh advfirewall export $wfwExportPath
WriteLog("End netsh advfirewall export")

netsh advfirewall firewall show rule name=all | Write-File -filePath $firewallFile -settingsString "netsh advfirewall firewall show rule"
Get-NetFirewallRule -All | Write-File -filePath $firewallFile -settingsString "Get-NetFirewallRule"

WriteLog("Start reg export FirewallPolicy")
reg export HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\SharedAccess\Parameters\FirewallPolicy $firewallRegistryFile
WriteLog("End reg export FirewallPolicy")

systeminfo | Write-File -filePath $systemInfoFile -settingsString "systeminfo"
Get-ComputerInfo | Write-File -filePath $systemInfoFile -settingsString "Get-ComputerInfo"

netstat -an | Write-File -filePath $networkFile -settingsString "netstat -an"
qwinsta /server:localhost | Write-File -filePath $networkFile -settingsString "qwinsta"
net share | Write-File -filePath $networkFile -settingsString "qwinsta"

net user | Write-File -filePath $userInfoFile -settingsString "net user"
query user | Write-File -filePath $userInfoFile -settingsString "query user"
query session | Write-File -filePath $userInfoFile -settingsString "query session"
net localgroup | Write-File -filePath $userInfoFile -settingsString "net localgroup"

tasklist | Write-File -filePath $taskInfoFile -settingsString "tasklist"
tasklist /svc | Write-File -filePath $taskInfoFile -settingsString "tasklist /svc"
qprocess | Write-File -filePath $taskInfoFile -settingsString "qprocess"

WriteLog("`r`nStart msinfo32 /report")
msinfo32 /report $msinfo32File
WriteLog("Start msinfo32 /report`r`n")


Get-WmiObject Win32_Product | Write-File -filePath $appInfoFile -settingsString "Get-WmiObject Win32_Product"
$apps = Get-ChildItem -Path(
  'HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
  'HKCU:SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall') | 
ForEach-Object {
  Get-ItemProperty $_.PsPath
}
$apps | Write-File -filePath $appInfoFile -settingsString "Apps"
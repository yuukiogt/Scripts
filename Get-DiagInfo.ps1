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

chcp 65001

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
$arpFile = Join-Path $logDir "Arp.log"
$firewallFile = Join-Path $logDir "Firewall.log"
$firewallRegistryFile = Join-Path $logDir "FirewallRegistry.log"
$systemInfoFile = Join-Path $logDir "SystemInfo.log"
$msinfo32File = Join-Path $logDir "msinfo32.log"
$userInfoFile = Join-Path $logDir "User.log"
$processInfoFile = Join-Path $logDir "Process.log"
$serviceInfoFile = Join-Path $logDir "Service.log"
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
netstat -an | Write-File -filePath $networkFile -settingsString "netstat -an"
qwinsta /server:localhost | Write-File -filePath $networkFile -settingsString "qwinsta"
net share | Write-File -filePath $networkFile -settingsString "net share"

WriteLog("Start netsh advfirewall export")
$wfwExportPath = (Join-Path $logDir "$($hostName)_firewall.wfw")
netsh advfirewall export $wfwExportPath
WriteLog("End netsh advfirewall export")

Get-NetFirewallRule -All | Write-File -filePath $firewallFile -settingsString "Get-NetFirewallRule"

WriteLog("Start reg export FirewallPolicy")
reg export HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\SharedAccess\Parameters\FirewallPolicy $firewallRegistryFile
WriteLog("End reg export FirewallPolicy")

systeminfo | Write-File -filePath $systemInfoFile -settingsString "systeminfo"
Get-ComputerInfo | Write-File -filePath $systemInfoFile -settingsString "Get-ComputerInfo"

$companyNetworkAdapter = Get-NetIPConfiguration | Where-Object { $_.NetProfile.Name -eq "nac社内ネットワーク" }
if ($Null -ne $companyNetworkAdapter) {
  $ipv4 = $companyNetworkAdapter.IPv4Address.IPAddress
  if($ipv4.StartsWith("172") -or $ipv4.StartsWith("192")) {
    $ipString = $ipv4.Split('.')
    $broadcastAddress = "$($ipString[0]).$($ipString[1]).0.255"
    ping $broadcastAddress | Write-File -filePath $arpFile -settingsString "ping $($broadcastAddress)"
    arp -a | Write-File -filePath $arpFile -settingsString "arp -a"
    net view | Write-File -filePath $arpFile -settingsString "net view"
  } else {
    WriteLog("unknown ipv4: $($ipv4)")
  }
}

net user | Write-File -filePath $userInfoFile -settingsString "net user"
query user | Write-File -filePath $userInfoFile -settingsString "query user"
query session | Write-File -filePath $userInfoFile -settingsString "query session"
net localgroup | Write-File -filePath $userInfoFile -settingsString "net localgroup"

Get-Process -FileVersionInfo | Select-Object ProductVersion, FileVersion, FileName, InternalName | Write-File -filePath $processInfoFile -settingsString "Get-Process"
qprocess | Write-File -filePath $processInfoFile -settingsString "qprocess"
Get-Service | Write-File -filePath $serviceInfoFile -settingsString "Get-Service"

WriteLog("Start msinfo32 /report")
msinfo32 /report $msinfo32File
WriteLog("End msinfo32 /report")

$apps = Get-ChildItem -Path(
  'HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
  'HKCU:SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall') | 
ForEach-Object {
  Get-ItemProperty $_.PsPath
}
$apps | Write-File -filePath $appInfoFile -settingsString "Apps"
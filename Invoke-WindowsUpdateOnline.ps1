$ConfirmPreference = "None"
$ErrorActionPreference = "SilentlyContinue"
$DebugPreference = "SilentlyContinue"
$InformationPreference = "SilentlyContinue"
$ProgressPreference = "SilentlyContinue"
$WarningPreference = "SilentlyContinue"
$VerbosePreference = "SilentlyContinue"

$currentDir = Split-Path $MyInvocation.MyCommand.Path -Parent
$scriptName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
$LogFile = Join-Path $currentDir ($scriptName + ".log")

function Log($message) {
    $now = Get-Date
    $log = $now.ToString("yyyy/MM/dd HH:mm:ss.fff") + "`t"
    $log += $message

    Write-Output $log | Out-File -FilePath $LogFile -Encoding UTF8 -append

    return $log
}

try {
    if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
    {
        $argsToAdminProcess = ""
        $Args.ForEach{ $argsToAdminProcess += "`"$PSItem`"" }
        Start-Process powershell.exe "-File `"$PSCommandPath`" $argsToAdminProcess" -Verb RunAs
        exit
    }

    Set-ItemProperty -Path 'HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate\AU' -Name 'UseWUServer' -Value 0
    Log("UseWUServer: 0")

    Restart-Service wuauserv
    Log("Restart-Service wuauserv")

    $updateSession = New-Object -ComObject Microsoft.Update.Session
    $searcher = $updateSession.CreateUpdateSearcher()
    $searchResult = $searcher.search("IsInstalled=0 and Type='software'")
    if ($searchResult.Updates.Count -eq 0) {
        Log("searchResult.Updates.Count 0")
    } else {
        $searchResult.Updates |
        ForEach-Object {
            $_.title -replace ".*(KB\d+).*", "`$1`t$&"
        }
    }

    $updatesToDownload = New-Object -ComObject Microsoft.Update.UpdateColl
    $searchResult.Updates | 
    Where-Object { -not $_.InstallationBehavior.CanRequestUserInput } |
    Where-Object { $_.EulaAccepted } |
    ForEach-Object {
        [void]$updatesToDownload.add($_)
        Log($_.title)
    }

    Log("Downloading ...")
    $downloader = $updateSession.CreateUpdateDownloader()
    $downloader.Updates = $updatesToDownload
    $downloader.Download()
    Log("Downloaded. ")

    $updatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
    $searchResult.Updates |
    Where-Object { $_.IsDownloaded } |
    ForEach-Object { [void]$updatesToInstall.add($_) }

    Log("Installing ... ")
    $installer = $updateSession.CreateUpdateInstaller()
    $installer.Updates = $updatesToInstall
    $installationResult = $installer.Install()
    Log("Installed. ")

    Log("installationResult: $($installationResult.ResultCode)")

} catch {
    Log($_.Exception.Message)
    pause
} finally {
    Set-ItemProperty -Path 'HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate\AU' -Name 'UseWUServer' -Value 1
    Log("UseWUServer: 1")

    Restart-Service wuauserv
    Log("Restart-Service wuauserv")
}
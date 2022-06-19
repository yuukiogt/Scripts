$ConfirmPreference = "None"
$ErrorActionPreference = "SilentlyContinue"
$DebugPreference = "SilentlyContinue"
$InformationPreference = "SilentlyContinue"
$ProgressPreference = "SilentlyContinue"
$WarningPreference = "SilentlyContinue"
$VerbosePreference = "SilentlyContinue"

$logDir = [Environment]::GetFolderPath('MyDocuments')
$scriptName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
$LogFile = Join-Path $logDir ($scriptName + ".log")

function WriteLog($message) {
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
        $Args.ForEach{ $argsToAdminProcess += " `"$PSItem`"" }
        Start-Process powershell.exe "-File `"$PSCommandPath`" $argsToAdminProcess"  -Verb RunAs
        exit
    }

    WriteLog("システムの復元ポイントを作成しています...")
    Enable-ComputerRestore -Drive C:\
    $checkpointDesc = Get-Date -Format yyyyMMddHHmmss
    Checkpoint-Computer -Description $checkpointDesc

    WriteLog("オンラインの Windows Update を On にしています...")
    Set-ItemProperty -Path 'HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate' -Name 'DisableWindowsUpdateAccess' -Value 0
    Set-ItemProperty -Path 'HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate\AU' -Name 'UseWUServer' -Value 0
    WriteLog("サービスを再起動しています...")
    Restart-Service wuauserv

    Read-Host @"
    一時的にオンラインの Windows Update を参照するように設定しました。

    このウィンドウを閉じずに、このまま Windows Update や オプション機能の追加を行ってください。

    完了後、何かキーを押してください。
"@

} catch {
    WriteLog($_.Exception.Message)
    pause
} finally {
    WriteLog("オンラインの Windows Update を Off にしています...")
    Set-ItemProperty -Path 'HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate\AU' -Name 'UseWUServer' -Value 1
    Set-ItemProperty -Path 'HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate' -Name 'DisableWindowsUpdateAccess' -Value 1
    WriteLog("サービスを再起動しています...")
    Restart-Service wuauserv

    Remove-Item $LogFile -Force
}

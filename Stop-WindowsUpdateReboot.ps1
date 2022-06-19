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

    $wuauclt = Get-Process | Where-Object ProcessName -eq "wuauclt"
    if($Null -eq $wuauclt) {
        Log("プロセスが存在しません")
    } else {
        Log("プロセスを停止します")
        Stop-Process -Id $wuauclt.Id
    }

    $wuauserv = Get-Service | Where-Object Name -eq "wuauserv"
    if ($Null -eq $wuauserv) {
        Log("サービスが存在しません")
    } else {
        Log("サービスを停止します")
        Stop-service -Name $wuauserv.Name
    }
}
catch {
    Log($_.Exception.Message)
    pause
}
finally {
    Remove-Item $LogFile -Force
}
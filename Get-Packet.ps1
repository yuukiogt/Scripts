if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $argsToAdminProcess = ""
    $Args.ForEach{ $argsToAdminProcess += "`"$PSItem`"" }
    Start-Process powershell.exe "-File `"$PSCommandPath`" $argsToAdminProcess" -Verb RunAs -WindowStyle Hidden
    exit
}

$sessionName = Read-Host "SessionName: "
$etlPath = "$($HOME)\Documents\Capture\$($sessionName).etl"

try {

    New-NetEventSession -Name $sessionName -LocalFilePath $etlPath

    (Get-NetEventProvider -ShowInstalled).Name | Select-String "TCP"

    Add-NetEventProvider -Name "Microsoft-Windows-TCPIP" -SessionName $sessionName

    Start-NetEventSession -Name $sessionName

    Start-Sleep -Seconds 1

    Get-NetEventSession

    Start-Sleep 10

    Stop-NetEventSession -Name $sessionName

    Start-Sleep -Seconds 1

    Get-NetEventSession
} catch {
    $_.Exception.Message
} finally {
    Remove-NetEventSession
    
    Start-Sleep -Seconds 1

    Get-NetEventSession
}

$session = Get-WinEvent -Path $etlPath -Oldest
$session[1..10]
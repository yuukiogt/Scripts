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

    $features = Get-WindowsCapability -Online |
    Where-Object {
        $_.Name.StartsWith("Language") -eq $False -and $_.Name.StartsWith("Rsat") -eq $False -and $_.State -ne "Installed"
    } |
    Select-Object Name

    $count = 0
    $featureSelectionStrings = $features |
    ForEach-Object {
        "$($count): $($_.Name)"
        $count++
    }

    $featureSelectionStrings += "q: 終了"
    $featureSelectionStrings += [System.Environment]::NewLine

    $auPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"

    $currentEncoding = $OutputEncoding
    $OutputEncoding = New-Object System.Text.ASCIIEncoding
    Enable-ComputerRestore -Drive C:\
    $checkpointDesc = Get-Date -Format yyyyMMddHHmmss
    Checkpoint-Computer -Description $checkpointDesc
    $OutputEncoding = $currentEncoding

    if($auPath) {
        Set-ItemProperty -LiteralPath $AUPath -Name "UseWUServer" -Value 0
        Restart-Service -Name "wuauserv"
        Write-Host "機能インストール前の処理を行なっています..."
    }

    do {
        $featureSelectionStrings
        $selectedNumber = Read-Host "インストールする機能の番号を選択してください 'q' を押すと終了します"

        if($selectedNumber -ge 0 -and $selectedNumber -lt $featureSelectionStrings.Length) {
            $featureName = $featureSelectionStrings[$selectedNumber].Split(" ")[1]
            Write-Host "$($featureName) をインストールしています..."
            Add-WindowsCapability -Online -Name $featureName
            Write-Host "$($featureName) をインストールしました"
        }
    } while($selectedNumber -ne 'q')

} catch {
    Write-Host $_.Exception.Message
    Log($_.Exception.Message)
} finally {
    if ($auPath) {
        Write-Host "機能インストール後の処理を行なっています..."
        Set-ItemProperty -LiteralPath $AUPath -Name "UseWUServer" -Value 1
        Restart-Service -Name "wuauserv"
    }
}
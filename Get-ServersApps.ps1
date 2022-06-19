$servers = Get-ADComputer -Filter { OperatingSystem -like "*Server*" } |
Where-Object { $_.Enabled -eq $True }

foreach ($server in $servers) {
    $computerName = $server.Name
    Write-Host "${computerName} のアプリ一覧を出力しています ..."
    Invoke-Command -ComputerName $computerName -ScriptBlock {
        $csvPath = ".\Apps\" + [Net.Dns]::GetHostName() + ".csv"
        Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* |
        Select-Object DisplayName, DisplayVersion, Publisher, InstallDate, InstallLocation, InstallSource, ModifyPath, HelpLink, UninstallString |
        Export-Csv -NoTypeInformation -Encoding UTF8 $csvPath

        Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* |
        Select-Object DisplayName, DisplayVersion, Publisher, InstallDate, InstallLocation, InstallSource, ModifyPath, HelpLink, UninstallString
    }
    Write-Host "${computerName} のアプリ一覧を出力しました"
}
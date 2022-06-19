$servers = @()

ForEach ($server in $servers) {
    if (!(Test-Connection $server -Count 1 -ErrorAction SilentlyContinue)) {
        continue
    }

    $target = "\\$($server)"
    psexec /s $target icacls "C:\PerfLogs" /t /q /grant "NAC\user:F"

    Move-Item -Path "$($target)\PerfLogs\System\Diagnostics\*" -Destination ".\Inventory\ServersSystemDiagReport\" -Force -Confirm:$false
}

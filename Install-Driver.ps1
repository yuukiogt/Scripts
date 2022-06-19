
function hasAdminAuth() {
    ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Start-ScriptAsAdmin([string]$scriptPath, [object[]]$argumentList) {
    if (!(hasAdminAuth)) {
        $list = @($scriptPath)
        if ($null -ne $argumentList) {
            $list += @($argumentList)
        }
        Start-Process powershell -ArgumentList $list -Verb RunAs -PassThru
    }
}

try {

    Start-ScriptAsAdmin -ScriptPath $PSCommandPath

    $path = Read-Host "Path "

    $drivers = Get-ChildItem $($path) -Recurse -Filter "*.inf"
    foreach($driver in $drivers) { 
        try {
            PNPUtil.exe /add-driver $driver.FullName /install
        }
        Catch {
            PNPUtil.exe /delete-driver $driver.Fullname /uninstall /force
            PNPUtil.exe /add-driver $driver.FullName /install
        }
    }
} catch {
    $_.Exception.Message
} finally {
    Read-Host "Press any key.."
}
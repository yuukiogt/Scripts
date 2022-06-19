$ErrorActionPreference = "Inquire"

function UpdateM365Installer([string]$path) {
    Set-Location -LiteralPath $path
    Start-Process -FilePath setup.exe -ArgumentList "/download Configuration.xml"
}

UpdateM365Installer -path ".\Path"
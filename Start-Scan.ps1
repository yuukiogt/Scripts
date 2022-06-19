$ConfirmPreference = "None"
$ErrorActionPreference = "SilentlyContinue"
$DebugPreference = "SilentlyContinue"
$InformationPreference = "SilentlyContinue"
$ProgressPreference = "SilentlyContinue"
$WarningPreference = "SilentlyContinue"
$VerbosePreference = "SilentlyContinue"

$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'

function Invoke-ExternalCommand([string]$commandPath, [string]$arguments) {
    try {
        $pinfo = New-Object System.Diagnostics.Process
        $pinfo.StartInfo.FileName = $commandPath
        $pinfo.StartInfo.Arguments = $arguments
        $pinfo.StartInfo.UseShellExecute = $false
        $pinfo.StartInfo.CreateNoWindow = $true
        $pinfo.StartInfo.UseShellExecute = $false
        $pinfo.StartInfo.RedirectStandardOutput = $true
        $pinfo.StartInfo.RedirectStandardError = $true
        $pinfo.StartInfo.CreateNoWindow = $true

        $oStdOutBuilder = New-Object -TypeName System.Text.StringBuilder
        $oStdErrBuilder = New-Object -TypeName System.Text.StringBuilder

        $sScripBlock = {
            if (! [String]::IsNullOrEmpty($EventArgs.Data)) {
                $Event.MessageData.AppendLine($EventArgs.Data)
            }
        }
        $oStdOutEvent = Register-ObjectEvent -InputObject $pinfo `
            -Action $sScripBlock -EventName 'OutputDataReceived' `
            -MessageData $oStdOutBuilder
        $oStdErrEvent = Register-ObjectEvent -InputObject $pinfo `
            -Action $sScripBlock -EventName 'ErrorDataReceived' `
            -MessageData $oStdErrBuilder

        $conhostsIDsBefore = (Get-Process -Name "conhost").Id
        [Void]$pinfo.Start()
        $conhostsIDsAfter = (Get-Process -Name "conhost").Id

        $targetID = $conhostsIDsAfter | Where-Object {$conhostsIDsBefore -notcontains $_ }

        $pinfo.BeginOutputReadLine()
        $pinfo.BeginErrorReadLine()
        
        $limit = 0
        while ($true) {
            Start-Sleep -Seconds 3
            $limit++
            if ($limit -gt 10) {
                break;
            }
            
            $outString = $oStdOutBuilder.ToString().Trim()
            if ($outString.IndexOf('結果を出力しました') -ne -1) {
                break;
            }
        }

        Unregister-Event -SourceIdentifier $oStdOutEvent.Name
        Unregister-Event -SourceIdentifier $oStdErrEvent.Name

        $oResult = New-Object -TypeName PSObject -Property ([Ordered]@{
                "ExitCode" = $pinfo.ExitCode;
                "stdout"   = $oStdOutBuilder.ToString().Trim();
                "stderr"   = $oStdErrBuilder.ToString().Trim()
            })
        return $oResult
    }
    finally {
        $pinfo.Kill()
        $pinfo.Dispose()
        Stop-Process -Id $targetID
    }
}

Add-Type -AssemblyName System.Windows.Forms

$emocheckx64 = "emocheck_x64.exe"
$emocheckx86 = "emocheck_x86.exe"

$tempPath = "C:\Temp\"
if (!(Test-Path $tempPath)) {
    exit
}

$dstx64 = Join-Path $tempPath $emocheckx64
$dstx86 = Join-Path $tempPath $emocheckx86

if ([System.Environment]::Is64BitProcess) {
    if (!$dstx64.EndsWith(".exe")) {
        exit
    }

    $oResult = Invoke-ExternalCommand $dstx64
}
else {
    if (!$dstx86.EndsWith(".exe")) {
        exit
    }

    $oResult = Invoke-ExternalCommand $dstx86
}

$PathTo = ".\destPath\"

$scanFiles = Get-ChildItem -LiteralPath $tempPath | Where-Object { $_ -like "*_emocheck.*" }
$scanFilesUser = Get-ChildItem -LiteralPath ([Environment]::GetFolderPath("UserProfile")) | Where-Object { $_ -like "*_emocheck.*" }
$scanFilesSystem = Get-ChildItem -LiteralPath ([Environment]::GetFolderPath("System")) | Where-Object { $_ -like "*_emocheck.*" }

if (Test-Path((Split-Path $PathTo -Parent))) {
    foreach ($file in $scanFiles) {
        Move-Item $file.FullName $PathTo -Force
    }

    foreach ($fileUser in $scanFilesUser) {
        Move-Item $fileUser.FullName $PathTo -Force
    }

    foreach ($fileSystem in $scanFilesSystem) {
        Move-Item $fileSystem.FullName $PathTo -Force
    }
}
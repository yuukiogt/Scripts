function Get-NumberOfLoggedOnUsers {
    $count = 0
    $error.Clear();

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo.FileName = "quser.exe"
    $process.StartInfo.UseShellExecute = $false
    $process.StartInfo.CreateNoWindow = $true
    $process.StartInfo.RedirectStandardOutput = $true
    $process.StartInfo.RedirectStandardError = $true
    $process.Start() | Out-Null

    $result = @()
    while ($null -ne ($line = $process.StandardOutput.ReadLine())) {
        if($line.Contains("Active")) {
            $result += $line
        }
    }

    if ([string]::isNullOrEmpty($process.StandardError.ReadToEnd()))
    {
        $count = $result.count
    }

    $process.WaitForExit()
    $process.Dispose()

    @{Count = $count}
}

(Get-NumberOfLoggedOnUsers).Count
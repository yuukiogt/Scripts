$DebugPreference = ”Stop”

$CurrentDir = Split-Path $MyInvocation.MyCommand.Path -Parent
$ScriptName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
$LogFile = Join-Path $CurrentDir ($ScriptName + ".log")

function WriteLog($message) {
    $now = Get-Date
    $log = $now.ToString("yyyy/MM/dd HH:mm:ss.fff") + "`t"
    $log += $message

    Write-Output $log | Out-File -FilePath $LogFile -Encoding UTF8 -append

    return $log
}

$returnCode = 0
$smart = Get-WmiObject -Namespace root\WMI -Class MSStorageDriver_FailurePredictData

ForEach($oClass in $smart)
{
    $checkID = @(1, 5, 9, 10, 13, 194, 196, 197, 198)

    WriteLog("Device = ($oClass.InstanceName)")

    for ($i = 2; $i -lt $oClass.VendorSpecific.Length; $i += 12) {
        if ($Null -ne $oClass.VendorSpecific[$i]) {
            if ($checkID -contains $oClass.VendorSpecific[$i]) {
                $rawValue = ””
                
                for ( $k = 8; $k -ge 5; $k-- ) {
                    $rawValue += [System.BitConverter]::ToString($oClass.VendorSpecific[($i + $k)], 0)
                }
                
                $rawValueInt = [Convert]::ToInt32($rawValue, 16)
                
                switch ($oClass.VendorSpecific[$i]) {
                    1 {
                        WriteLog("Raw Read Error Rate")
                        if ( $rawValueInt -gt 0 ) {
                            $returnCode = 2
                        }
                        $checkID[0] = 9999
                    }
                    5 {
                        WriteLog("Reallocated Sectors Count")
                        if ( $rawValueInt -gt 0 ) {
                            $returnCode = 2
                        }
                        $checkID[1] = 9999
                    }
                    9 {
                        WriteLog("Power-On Hours")
                        $checkID[2] = 9999
                    }
                    10 {
                        WriteLog("Spin Retry Count")
                        if ( $rawValueInt -gt 0 ) {
                            $returnCode = 2
                        }
                        $checkID[3] = 9999
                    }
                    13 {
                        WriteLog("Soft Read Error Rate")
                        if ( $rawValueInt -gt 0 ) {
                            $returnCode = 2
                        }
                        $checkID[4] = 9999
                    }
                    194 {
                        WriteLog("Temperature")
                        if ( $rawValueInt -gt 55 ) {
                            $returnCode = 2
                        }
                        $checkID[5] = 9999
                    }
                    196 {
                        WriteLog("Reallocation Event Count")
                        if ( $rawValueInt -gt 0 ) {
                            $returnCode = 2
                        }
                        $checkID[6] = 9999
                    }
                    197 {
                        WriteLog("Current Pending Sector Count")
                        if ( $rawValueInt -gt 0 ) {
                            $returnCode = 2
                        }
                        $checkID[7] = 9999
                    }
                    198 {
                        WriteLog("Off-Line Scan Uncorrectable Sector Count")
                        if ( $rawValueInt -gt 0 ) {
                            $returnCode = 2
                        }
                        $checkID[8] = 9999
                    }
                }

                for($j = $i; $j -le $i + 11; $j++) {
                    if ( $Null -ne $oClass.VendorSpecific[$j] ) {
                        WriteLog("$oClass.VendorSpecific[$j].ToString()")
                    }
                }
                WriteLog("Total = $rawValueInt")
            }
        }
    }
}

if( $returnCode -ne 0) {

}

exit $returnCode
chcp 65001
$profiles = (netsh wlan show profiles)

$ssid = ": "

for($i = 0; $i -lt $profiles.length; $i++)
{
    if($profiles[$i] -match $ssid)
    {
        $lanName=$profiles[$i].Split(":")[1].Split(" ")
        Write-Output $lanName[1]
        $result=(netsh wlan connect name=($lanName[1]))
        Write-Output $result
    }
}

[Windows.Devices.WiFi.WiFiAdapter, Windows.Devices.WiFi, ContentType = WindowsRuntime] | Out-Null
Add-Type -AssemblyName System.Runtime.WindowsRuntime
Add-Type -AssemblyName System.Runtime.InteropServices.WindowsRuntime
 
Function Wait-IAsyncOperation {
    param(
        [Parameter(Mandatory)]
        [System.Management.Automation.PSMethod] $Method,
        [Parameter(Mandatory = $false)]
        [object[]] $Arguments = @(),
        [Parameter(Mandatory = $false)]
        [type] $ResultType
    )
    if (! $ResultType) {
        if ($Method.OverloadDefinitions[0] -match 'IAsyncOperation\[([\w\.\[\]]+)\] ') {
            $ResultType = $Matches[1] -as [type]
            Write-Verbose $ResultType
        }
        else {
            Write-Warning "不正なメソッド: $Method"
            break
        }
    }
    $asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
    $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
    $netTask = $asTask.Invoke($null, @($Method.Invoke($Arguments)))
    try {
        $netTask.Wait(-1) | Out-Null
        $netTask.Result
    }
    catch {
        Write-Error $netTask.Exception
    }
}
 
function Get-WifiAdapter {
    [OutputType([Windows.Devices.WiFi.WiFiAdapter])]
    [CmdletBinding()]
    param(
 
    )
    process {
        $selector = [Windows.Devices.WiFi.WiFiAdapter]::GetDeviceSelector()
        Wait-IAsyncOperation -Method ([Windows.Devices.WiFi.WiFiAdapter]::FindAllAdaptersAsync)
    }
}
 
function Get-WiFiAvailableNetwork {
    [OutputType([Windows.Devices.WiFi.WiFiAvailableNetwork])]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [Windows.Devices.WiFi.WiFiAdapter] $WifiAdapter,
        [Parameter()]
        [string] $Ssid = '*'
    )
    process {
        $WifiAdapter.NetworkReport.AvailableNetworks | ? Ssid -Like $Ssid
    }
}
 
function Disconnect-WifiAdapter {
    [OutputType()]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [Windows.Devices.WiFi.WiFiAdapter] $WifiAdapter
    )
    process {
        $WifiAdapter.Disconnect()
    }
}
 
function Connect-WifiAdapter {
    [OutputType([Windows.Devices.WiFi.WiFiConnectionResult])]
    [CmdletBinding(DefaultParameterSetName = 'pipenetwork')]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'pipeadapter')]
        [Parameter(Mandatory, ParameterSetName = 'pipenetwork')]
        [Windows.Devices.WiFi.WiFiAdapter] $WifiAdapter,
        [Parameter(Mandatory, ParameterSetName = 'pipeadapter')]
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'pipenetwork')]
        [Windows.Devices.WiFi.WiFiAvailableNetwork] $WiFiAvailableNetwork,
        [Parameter(Mandatory = $false)]
        [Windows.Devices.WiFi.WiFiReconnectionKind] $WiFiReconnectionKind = 'Manual',
        [Parameter(Mandatory = $false)]
        [string] $Password,
        [Parameter(Mandatory = $false)]
        [string] $HiddenSsid
    )
    process {
        $args = @(
            $WiFiAvailableNetwork
            $WiFiReconnectionKind
        )
        if ($Password) {
            $args += [Windows.Security.Credentials.PasswordCredential]::new('tmp', 'tmp', $Password)
            if ($HiddenSsid) {
                $args += $HiddenSsid
            }
        }
        Wait-IAsyncOperation -Method ($WifiAdapter.ConnectAsync) -Arguments $args
    }
}
 
Export-ModuleMember -Function *-Wifi*

$adapter = Get-WifiAdapter
$ap = Get-WiFiAvailableNetwork -WifiAdapter $adapter -Ssid "YourSsid"
Connect-WifiAdapter -WifiAdapter $adapter -WiFiAvailableNetwork $ap -WiFiReconnectionKind Automatic

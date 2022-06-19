#Use the following syntax to get the local computer's UUID/GUID using Windows Powershell:
$guid = Get-Wmiobject Win32_ComputerSystemProduct | Select-Object -ExpandProperty UUID
$guid

#Add -computername after the WMI class to find a remote computer's UUID, example:
$uuid = Get-Wmiobject Win32_ComputerSystemProduct -Computername ([Net.Dns]::GetHostName()) | Select-Object -ExpandProperty UUID
$uuid
$isLaptop = [bool](Get-CimInstance -ClassName Win32_SystemEnclosure).ChassisTypes.Where({ $PSItem -in 9, 10, 14 })

$isLaptop
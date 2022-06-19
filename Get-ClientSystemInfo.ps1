$target = Read-Host "ComputerName"

Invoke-Command -ComputerName $target -ScriptBlock {
    Get-ComputerInfo
}
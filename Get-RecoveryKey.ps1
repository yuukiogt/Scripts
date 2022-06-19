$drives = Get-WmiObject Win32_LogicalDisk | ForEach-Object { $_.Name }

foreach ($drive in $drives) {    
    $result = manage-bde -protectors -get $drive
    Write-Host $result
}

Read-Host "Enterを押すと終了します"
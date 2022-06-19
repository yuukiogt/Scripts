$computers = Get-ADComputer -Filter { OperatingSystem -Like '*Windows 7*' -and Enabled -eq $True }

$credential = Get-Credential

foreach ($computer in $computers) { 
    $pcinfo = Get-ADComputer $computer.Name -Properties lastlogontimestamp |
    Select-Object @{Name = "Computer"; Expression = { $_.Name } }, @{Name = "Lastlogon"; Expression = { [DateTime]::FromFileTime($_.lastLogonTimestamp) } }

    $lastuserlogoninfo = Get-WmiObject -Class Win32_UserProfile -ComputerName $computer.name -Credential $credential |
    Select-Object -First 1
    $SecIdentifier = New-Object System.Security.Principal.SecurityIdentifier($lastuserlogoninfo.SID)
    $username = $SecIdentifier.Translate([System.Security.Principal.NTAccount])

    $properties = @{
        'Computer' = $pcinfo.Computer;
        'LastLogon' = $pcinfo.Lastlogon;
        'User' = $username.value
    }

    Write-Output (New-Object -Typename PSObject -Property $properties)
}
# https://docs.microsoft.com/ja-jp/windows-server/manage/windows-admin-center/use/get-started

$Computers = Get-ADComputer -Filter { Enabled -eq $True } -Properties * |
Select-Object Name,DNSHostName, OperatingSystem

$ServerCsvPath = ".\ServerNamesForWindowsAdminCenter.csv"
$isExistCSV = Test-Path $ServerCsvPath

if ($isExistCSV -eq $True) {
    Remove-Item $ServerCsvPath -Force
}

$ComputerCsvPath = ".\ComputerNamesForWindowsAdminCenter.csv"
$isExistCSV = Test-Path $ComputerCsvPath

if ($isExistCSV -eq $True) {
    Remove-Item $ComputerCsvPath -Force
}

foreach($Computer in $Computers) {
    $DNSHostName = $Computer.DNSHostName
    if([string]::IsNullOrEmpty($DNSHostName)) {
        $DNSHostName = $Computer.Name
    }

    if($Computer.OperatingSystem -like "*Server*") {
        Add-Content -Path $ServerCsvPath -Value "${DNSHostName}"
    } else {
        Add-Content -Path $ComputerCsvPath -Value "${DNSHostName}"
    }
}
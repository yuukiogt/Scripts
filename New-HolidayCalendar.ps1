$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $UserCredential -Authentication "Basic" -AllowRedirection
Import-PSSession $Session -DisableNameChecking

[void][Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.Outlook")

$olDefaultFolders = [Microsoft.Office.Interop.Outlook.OlDefaultFolders]
$olItemType = [Microsoft.Office.Interop.Outlook.OlItemType]
$olBusyStatus = [Microsoft.Office.Interop.Outlook.OlBusyStatus]

$folder = $namespace.GetDefaultFolder($OlDefaultFolders::olFolderCalendar)

$date = [DateTime]::ParseExact("20200101", "yyyyMMdd", $null)

$newItem = $outlook.CreateItem($OlItemType::olAppointmentItem)
$newItem.Subject = "終日の予定"
$newItem.Start = $date
$newItem.End = $date.AddDays(1)
$newItem.AllDayEvent = $true
$newItem.BusyStatus = $OlBusyStatus::olFree
$newItem.Save()
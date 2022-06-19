$credential = Get-Credential

$siteCollectionUrl = "https://tenant.sharepoint.com/sites/nic"
$connection = Connect-PnPOnline -Url $siteCollectionUrl -Credentials ($credential)

Get-PnPList -Connection $connection
$listName = "ComputerEvent"

$deleteItem = Get-PnPListItem -List $listName -Fields "FileRef" | Where-Object { $_.FieldValues['FileRef'].EndsWith("Event_v2.csv") }
if($Null -ne $deletedItem) {
    Remove-PnPListItem -List $listName -Identity $deleteItem.Id -Force
}

$items = Get-PnPListItem -List $listName -Fields "FileRef"

[string]$streamString = "HostName,UserName,DateTime,EventId,Description`r`n"

foreach($item in $items) {
    $fileUrl = $item.FieldValues['FileRef']
    $file = Get-PnPFile -Url $fileUrl -AsString
    $lines = $file.Split("`n")

    $dateSpans = @{}

    for ($i = 1; $i -lt $lines.Length; $i++) {
        $datetimeString = $lines[$i].Split(',')[2]
        if($Null -ne $dateTimeString) {
            $targetDate = $datetimeString.Split(" ")[0]
            if ($dateSpans.ContainsKey($targetDate)) {
                $dateSpans[$targetDate] += $lines[$i]
            }
            else {
                $dateSpans[$targetDate] = @($lines[$i])
            }
        }
    }

    $dateSpans = $dateSpans.GetEnumerator() | Sort-Object -Property Key

    # 6005 { $Description = "スタートアップ" }
    # 6006 { $Description = "シャットダウン" }
    # 6008 { $Description = "異常終了" }
    # 7001 { $Description = "サインイン" }
    # 7002 { $Description = "サインアウト" }
    # 20267 { $Description = "VPN 接続" }
    # 20268 { $Description = "VPN 切断" }

    $startTimeString = [string]::Empty
    $endTimeString = [string]::Empty
    $startCurrentString = [string]::Empty
    $endCurrentString = [string]::Empty
    
    foreach($datespan in $dateSpans) {
        foreach($value in $datespan.Value) {
            $eventId = $value.Split(",")[3]
            if ( $eventId -eq 6006 -or $eventId -eq 7002 ) {
                if (![string]::IsNullOrEmpty($value)) {
                    $endTimeString = $value.Split(",")[2]
                    if(![string]::IsNullOrEmpty($endTimeString)) {
                        $endTimeString = $endTimeString.Substring(0, $endTimeString.Length - 3)
                    }
                    $endCurrentString = $value
                    $streamString += $value
                    break
                }
            }
        }

        [array]::Reverse($datespan.Value)
        foreach ($value in $datespan.Value) {
            $eventId = $value.Split(",")[3]
            if ($eventId -eq 6005 -or $eventId -eq 7001 ) {
                if (![string]::IsNullOrEmpty($value)) {
                    $startTimeString = $value.Split(",")[2]
                    if(![string]::IsNullOrEmpty($startTimeString)) {
                        $startTimeString = $startTimeString.Substring(0, $startTimeString.Length - 3)
                    }
                    $startCurrentString = $value
                    $streamString += $value
                    break
                }
            }
        }

        if (![string]::IsNullOrEmpty($streamString) -and $startTimeString -eq $endTimeString) {
            $stringCount = $startCurrentString.Length + $endCurrentString.Length
            $streamString = $streamString.Substring(0,$streamString.Length - $stringCount)
            $streamString += $startCurrentString
        }
    }
}

$tempFile = Join-Path $env:TEMP "ComputerEvent_v2.csv"
New-Item -ItemType File -Path $tempFile

$streamWriter = New-Object System.IO.StreamWriter($tempFile, $false, [Text.Encoding]::GetEncoding("UTF-8"))
$streamWriter.WriteLine($streamString)
$streamWriter.Dispose()

$tempFile = Get-Content -Path $tempFile | Sort-Object -Property DateTime -Unique

Add-PnPFile -Path $tempFile -Folder "ComputerEvent"

Remove-Item -Path $tempFile -Force

Disconnect-PnPOnline -Connection $connection

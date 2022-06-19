$SiteURL = "https://tenant.sharepoint.com/sites/automotive"
$CSVFile = ".\VersionHistoryReport.csv"

If (Test-Path $CSVFile) {
    Remove-Item $CSVFile
}

$connect = Connect-PnPOnline -Url $SiteURL -Credentials (Get-Credential) # -UseWebLogin

try {
$ExcludedLists = @("Form Templates", "Preservation Hold Library", "Site Assets", "Pages", "Site Pages", "Images",
    "Site Collection Documents", "Site Collection Images", "Style Library")
$Lists = Get-PnPList | Where-Object { $_.Hidden -eq $False -and $_.Title -notin $ExcludedLists -and $_.BaseType -eq "DocumentLibrary" }
 
ForEach ($List in $Lists) {
    $global:counter = 0
    $Files = Get-PnPListItem -List $List -PageSize 2000 -Fields File_x0020_Size, FileRef -ScriptBlock {
        Param($items) $global:counter += $items.Count; Write-Progress -PercentComplete ($global:Counter / ($List.ItemCount) * 100) -Activity "Getting Files of '$($List.Title)'" -Status "Processing Files $global:Counter to $($List.ItemCount)";
    }  | Where-Object { $_.FileSystemObjectType -eq "File"}
     
    $VersionHistoryData = @()
    $Files | ForEach-Object {
        Write-host "Getting Versioning Data of the File:"$_.FieldValues.FileRef

        $FileSizeinKB = [Math]::Round(($_.FieldValues.File_x0020_Size / 1KB), 2)
        $File = Get-PnPProperty -ClientObject $_ -Property File
        $Versions = Get-PnPProperty -ClientObject $File -Property Versions
        $VersionSize = $Versions | Measure-Object -Property Size -Sum | Select-Object -expand Sum
        $VersionSizeinKB = [Math]::Round(($VersionSize / 1KB), 2)
        $TotalFileSizeKB = [Math]::Round(($FileSizeinKB + $VersionSizeinKB), 2)
  
        $VersionHistoryData += New-Object PSObject -Property  ([Ordered]@{
                "Library Name"         = $List.Title
                "File Name"            = $_.FieldValues.FileLeafRef
                "File URL"             = $_.FieldValues.FileRef
                "Versions"             = $Versions.Count
                "File Size (KB)"       = $FileSizeinKB
                "Version Size (KB)"    = $VersionSizeinKB
                "Total File Size (KB)" = $TotalFileSizeKB
            })
    }
    $VersionHistoryData | Export-Csv -Path $CSVFile -NoTypeInformation -Append -Encoding UTF8
}
} catch {
    $_.Exception.Message
} finally {
    if($connect) {
        Disconnect-PnPOnline
    }
}
$baseFile = Import-Csv ".\base.csv" -Encoding UTF8 | Select-Object DisplayName
$compFile = Import-Csv ".\comp.csv" -Encoding UTF8 | Select-Object Name

$delimiters = @('','','')
$baseFile
$compFile

$diffBaseData = @()
$diffCompData = @()
$matchData = @()

if (($diffResult -ne $null) -and ($diffResult.Count -gt 0)) {
    foreach ($data in $diffResult) {
        $indicator = $data.SideIndicator
        $inputObject = $data.InputObject

        if ($indicator -eq "<=") {
            $diffBaseData += $inputObject
        }
        elseif ($indicator -eq "=>") {
            $diffCompData += $inputObject
        }
        elseif ($indicator -eq "==") {
            $matchData += $inputObject
        }
        else {
            Write-Host "Invalid Line"
        }
    }

    $diffBaseCsv = $diffBaseData | ConvertFrom-Csv -Header "Name", "Age", "Email"
    $diffCompCsv = $diffCompData | ConvertFrom-Csv -Header "Name", "Age", "Email"

    $matchCsv = $matchData    | ConvertFrom-Csv

    $diffBaseCsv | Export-Csv -NoTypeInformation .\DiffBaseFile.csv -Encoding UTF8
    $diffCompCsv | Export-Csv -NoTypeInformation .\DiffCompFile.csv -Encoding UTF8
    $matchCsv    | Export-Csv -NoTypeInformation .\MatchFile.csv -Encoding UTF8
}
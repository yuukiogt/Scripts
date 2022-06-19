$ErrorActionPreference = "Inquire"

$CurrentDir = Split-Path $MyInvocation.MyCommand.Path -Parent
$ScriptName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
$LogFile = Join-Path $CurrentDir ($ScriptName + ".log")

$BodyTempFilePath = "BodyTemp.xlsx"
$Department = "*IT*"

function Log($message) {
    $now = Get-Date
    $log = $now.ToString("yyyy/MM/dd HH:mm:ss.fff") + "`t"
    $log += $message

    Write-Output $log | Out-File -FilePath $LogFile -Encoding UTF8 -append

    return $log
}

try {
    $Excel = New-Object -ComObject Excel.Application
    $Workbook = $Excel.Workbooks.Open($BodyTempFilePath)
    $Excel.Visible = $False

    do {
        $TargetMonth = $Null
        $IsExistMonth = $True

        do {
            [int]$TargetDateString = Read-Host "月を入力してください 1～12"
            if ($TargetDateString -ge 1 -and $TargetDateString -le 12) {
                $TargetMonth = Get-Date -Month $TargetDateString -Day 1
            }
            if ($TargetDateString -eq 1) {
                $TargetMonth = $TargetMonth.AddMonths(12)
            }
        } while ($Null -eq $TargetMonth)

        $TargetMonthString = $TargetMonth.ToString("yyyy-MM")

        try {
            $Workbook.Worksheets.Item($TargetMonthString)
            Write-Host "指定した月のシートは既に存在しています..."
        } catch {
            $IsExistMonth = $False
        }
    } while ($IsExistMonth -eq $True)

    $ActiveBook = $Excel.ActiveWorkbook
    $Workbook.WorkSheets.Add([System.Reflection.Missing]::Value, $ActiveBook.Sheets($ActiveBook.Sheets.Count))

    $Workbook.Worksheets.Item("Sheet1").Name = $TargetMonthString

    $TargetWorksheet = $Workbook.Worksheets.Item($TargetMonthString)
    $TargetWorksheet.Cells(1,1) = "名前"

    $LastDay = [int]$TargetMonth.AddMonths(1).AddDays(-1).ToString("dd")
    $DayIndex = 1
    [int]$MonthIndex = $TargetMonth.ToString("MM")
    For($ColumnIndex = 2; $ColumnIndex -le $LastDay + 1; $ColumnIndex++) {
        $TargetWorksheet.Cells(1, $ColumnIndex) = (Get-Date -Year $TargetMonth.Year -Month $MonthIndex -Day $DayIndex).ToString("d (ddd)")
        $DayIndex++
    }

    $Users = Get-ADUser -Properties * -Filter { Department -like $Department } | Select-Object Name, msDS-PhoneticDisplayName
    $Users = $Users | Sort-Object msDS-PhoneticDisplayName

    $NameIndex = 2
    foreach($User in $Users) {
        $TargetWorksheet.Cells($NameIndex, 1) = $User.Name
        $NameIndex++
    }

    $TargetWorksheet.Range($TargetWorksheet.Cells.Item(2, 2), $TargetWorksheet.Cells.Item($Users.Count + 1, $LastDay + 1)).NumberFormat = "00.0"

    $ListObject = $Excel.ActiveSheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $Excel.ActiveCell.CurrentRegion, $null , [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
    $ListObject.TableStyle = "TableStyleLight1"

    $Workbook.Save()

} catch {
    $_.Exception.Message
} finally {
    $ActiveBook = $Null
    $Workbook.Close($False)
    $Workbook = $Null
    $Excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($TargetWorksheet)
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Excel)
}
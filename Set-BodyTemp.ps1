$ErrorActionPreference = "Inquire"

$currentDir = Split-Path $MyInvocation.MyCommand.Path -Parent
$scriptName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
$logFile = Join-Path $currentDir ($scriptName + ".log")

function Log($message) {
    $now = Get-Date
    $log = $now.ToString("yyyy/MM/dd HH:mm:ss.fff") + "`t"
    $log += $message

    Write-Output $log | Out-File -FilePath $logFile -Encoding UTF8 -append

    return $log
}

try {
    if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
    {
        $argsToAdminProcess = ""
        $Args.ForEach{ $argsToAdminProcess += "`"$PSItem`"" }
        Start-Process powershell.exe "-File `"$PSCommandPath`" $argsToAdminProcess" -Verb RunAs
        exit
    }

    $targetMonth = Get-Date -Format yyyy-MM
    Write-Host "TargetMonth: $($targetMonth)"

    $bodyTempFilePath = ".\体温表.xlsx"
    Write-Host "File: $($bodyTempFilePath)"

    $excel = New-Object -ComObject Excel.Application
    $workbook = $excel.Workbooks.Open($bodyTempFilePath)
    $excel.Visible = $False
    Write-Host "Workbooks.Open"

    $sheet = $workbook.Worksheets.Item($targetMonth)
    $targetDate = (Get-Date).ToString("d (ddd)")
    Write-Host "TargetDate: $($targetDate)"

    $targetCellColumn = $sheet.Cells.Find($targetDate).Column
    $targetCellRow = $sheet.Cells.Find("MyName").Row
    Write-Host "Cell: ($($targetCellRow),$($targetCellColumn))"

    $temp = {
        $randomValue = Get-Random -Maximum 36.7 -Minimum 36.2
        $randomValue = [Math]::Round($randomValue, 1, [MidpointRounding]::AwayFromZero)
        Write-Host "Temp: $($randomValue)"
        return $randomValue
    }

    $tempValue = & $temp
    $sheet.Cells.Item($targetCellRow, $targetCellColumn) = $tempValue

    $count = 1
    while ($True) {
        $prevCellColumn = $targetCellColumn - $count
        $prevCell = $sheet.Cells.Item($targetCellRow, $prevCellColumn)
        if ($prevCell.Text -eq [string]::Empty) {
            Write-Host "prevCell: ($($targetCellRow), $($prevCellColumn))"

            $tempValue = & $temp
            $sheet.Cells.Item($targetCellRow, $prevCellColumn) = $tempValue
        }
        else {
            break;
        }
        $count++
    }

    $workbook.Save()

    Write-Host "Excel.Save"
} catch {
    Write-Host $_.Exception.Message
    Log($_.Exception.Message)
} finally {
    $workbook.Close()
    Write-Host "Workbook.Close"

    $excel.Quit()
    Write-Host "Excel.Quit"

    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook)
    Write-Host "FinalReleaseComObject.Workbook"

    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    Write-Host "FinalReleaseComObject.Excel"

    [GC]::Collect()
    Write-Host "GC.Collect"

    Get-Process -Name Excel |
    ForEach-Object {
        Stop-Process -id $_.Id
        Write-Host "Excel Stop-Process ID: $($_.Id)"
    }
}
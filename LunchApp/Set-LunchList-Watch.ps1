$CurrentDir = Split-Path $MyInvocation.MyCommand.Path -Parent
$ScriptName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
$LogFile = Join-Path $CurrentDir ($ScriptName + ".log")

function Log($message) {
    $now = Get-Date
    $log = $now.ToString("yyyy/MM/dd HH:mm:ss.fff") + "`t"
    $log += $message

    Write-Output $log | Out-File -FilePath $LogFile -Encoding UTF8 -append

    return $log
}

try {
    $M365Credential = Get-Credential
    Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
    Connect-SPOService -Url https://tenant-admin.sharepoint.com -Credential $M365Credential | Out-Null

    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client") | Out-Null
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime") | Out-Null

    $SiteUrl = "https://tenant.sharepoint.com/sites/site/"
    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
    $SPCredential = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName, $SecurePassword)
    $Context.Credentials = $SPCredential
    $WebObject = $Context.Web
    $Context.Load($WebObject)
    $Context.ExecuteQuery()

    $ListName = "お弁当注文_" + (Get-Date).AddMonths(1).ToString("yyyy")
    $List = $Context.Web.Lists.GetByTitle($ListName)
    $Context.Load($List)
    $Context.ExecuteQuery()

    $SPOFile = $WebObject.ServerRelativeUrl + "/Lunch/Lunch.xlsx"
    $ToFile = Join-Path $CurrentDir "Lunch.xlsx"

    $OpenFile = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($Context, $SPOFile)
    $WriteStream = [System.IO.File]::Open($ToFile, [System.IO.FileMode]::Create)
    $OpenFile.Stream.CopyTo($WriteStream)

    $WriteStream.Close()

    $Excel = New-Object -ComObject Excel.Application
    $Workbook = $Excel.Workbooks.Open($ToFile)
    $Worksheet = $Workbook.Sheets.Item("Menu")
    $Excel.Visible = $False

    $Columns = $WorkSheet.UsedRange.Columns.Count
    $Rows = $WorkSheet.UsedRange.Rows.Count

    $DeadlineFieldIndex = 1
    For ($Column = 1; $Column -le $Columns; $Column++) {
        $FieldName = $WorkSheet.Columns.Item($Column).Rows.Item(1).Text
        if($FieldName -eq "Deadline") {
            $DeadlineFieldIndex = $Column
            break
        }
    }

    $TargetDateFlag = $False
    $TargetDate = $Null

    For ($Row = 2; $Row -le $Rows; $Row++) {
        if ($WorkSheet.Columns.Item($DeadlineFieldIndex).Rows.Item($Row).Text.Contains((Get-Date).ToString("yyyy/M/dd"))) {
            $TargetDateFlag = $True
            continue
        }
        if($True -eq $TargetDateFlag) {
            $TargetDate = [System.DateTime]($WorkSheet.Columns.Item($DeadlineFieldIndex).Rows.Item($Row).Text)
            break
        }
    }
    Log("TargetDate: $TargetDate")

    $Rows = $Null
    $Columns = $Null
    $Worksheet = $Null
    $Workbook.Close($False)
    $Workbook = $Null
    $Excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
    Remove-Variable Excel

    if($Null -ne $TargetDate) {
        $Fields = $List.Fields
        $Context.Load($Fields)
        $Context.ExecuteQuery()

        $Field = $List.Fields | Where-Object { $_.Title -eq ($TargetDate.ToString("M/d (ddd)")) }

        if($Null -ne $Field) {
            $Field.ReadOnlyField = $True
            $Field.Update()
            $Context.ExecuteQuery()
        }
    }

    $Context.Dispose()
    Disconnect-SPOService
} catch {
    $ErrorMessage = $_.Exception.Message
    Log($ErrorMessage)
}
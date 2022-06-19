$CurrentDir = Split-Path $MyInvocation.MyCommand.Path -Parent
$PasswordFile = Join-Path $CurrentDir "Password.txt"
$ScriptName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
$LogFile = Join-Path $CurrentDir ($ScriptName + ".log")

function Log($message) {
    $now = Get-Date
    $log = $now.ToString("yyyy/MM/dd HH:mm:ss.fff") + "`t"
    $log += $message

    Write-Output $log | Out-File -FilePath $LogFile -Encoding UTF8 -append

    return $log
}

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
Log("ClientContext Executed")

$TargetMonth = $Null

do {
    [int]$TargetDateString = Read-Host "追加する月を入力してください 1～12"
    if ($TargetDateString -ge 1 -and $TargetDateString -le 12) {
        $TargetMonth = Get-Date -Month $TargetDateString -Day 1
    }
    if ($TargetDateString -eq 1) {
        $TargetMonth = $TargetMonth.AddMonths(12)
    }
} while ($Null -eq $TargetMonth)

$ListName = "お弁当注文_" + (Get-Date).AddMonths(1).ToString("yyyy")
$List = $Context.Web.Lists.GetByTitle($ListName)
$Context.Load($List)
$Context.ExecuteQuery()
Log("$ListName GetByTitle Executed")

$Fields = $List.Fields
$Context.Load($Fields)
$Context.ExecuteQuery()
Log("Fields Load Executed")

$SPOFile = $WebObject.ServerRelativeUrl + "/Lunch/Lunch.xlsx"
$ToFile = Join-Path $CurrentDir "Lunch.xlsx"

$OpenFile = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($Context, $SPOFile)
$WriteStream = [System.IO.File]::Open($ToFile, [System.IO.FileMode]::Create)
$OpenFile.Stream.CopyTo($WriteStream)
Log("CopyTo $ToFile")

$WriteStream.Close()

$Excel = New-Object -ComObject Excel.Application
$Workbook = $Excel.Workbooks.Open($ToFile)
$Worksheet = $Workbook.Sheets.Item("Menu")
$Excel.Visible = $False

$Columns = $WorkSheet.UsedRange.Columns.Count
$Rows = $WorkSheet.UsedRange.Rows.Count

$MenuFields = @()
For ($Column = 1; $Column -le $Columns; $Column++) {
    $FieldName = $WorkSheet.Columns.Item($Column).Rows.Item(1).Text
    $MenuFields += $FieldName
}

$DataCollection = @()
For ($Row = 2; $Row -le $Rows; $Row++) {
    if ($WorkSheet.Columns.Item(1).Rows.Item($Row).Text.Contains($TargetMonth.ToString("yyyy/M"))) {
        $Data = New-Object PSObject
        For ($Column = 1; $Column -le $Columns; $Column++) {
            $Value = $WorkSheet.Columns.Item($Column).Rows.Item($Row).Text
            $Data | Add-Member -MemberType noteproperty -Name $MenuFields[$Column - 1] -Value $Value
        }
        $DataCollection += $Data
    }
}

$Worksheet = $Workbook.Sheets.Item("Value")
$Columns = $WorkSheet.UsedRange.Columns.Count
$Rows = $WorkSheet.UsedRange.Rows.Count

$ValueFields = @()
For ($Column = 1; $Column -le $Columns; $Column++) {
    $FieldName = $WorkSheet.Columns.Item($Column).Rows.Item(1).Text
    $ValueFields += $FieldName
}

$MenuValues = [ordered]@{}
For ($Row = 2; $Row -le $Rows; $Row++) {
    $Key = $WorkSheet.Columns.Item(1).Rows.Item($Row).Text
    $Value = [int]($WorkSheet.Columns.Item(2).Rows.Item($Row).Text)
    $MenuValues.Add($Key, $Value)
}

$Rows = $Null
$Columns = $Null
$Worksheet = $Null
$Workbook.Close($False)
$Workbook = $Null
$Excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
Remove-Variable Excel

Log("Excel.Quit()")

$InitialDate = Get-Date -Date $DataCollection[0].Date
$TargetDate = $InitialDate
$LastDate = Get-Date -Date $DataCollection[$DataCollection.Length - 1].Date

ForEach ($Data in $DataCollection) {
    $TargetDate = [System.DateTime]$Data.Date
    $FieldName = "Day" + $TargetDate.ToString("yyyyMMdd")
    $Field = $List.Fields | Where-Object { $_.InternalName -eq $FieldName }
    if ($Null -ne $Field) {
        if($True -eq $Field.ReadOnlyField) {
            $Field.ReadOnlyField = $False
            $Field.Update()
            $Context.ExecuteQuery()
        }
        $Field.DeleteObject()
        $Context.ExecuteQuery()
    }
}

$TargetDate = $InitialDate

$TargetMenuValues = [ordered]@{}

ForEach($Data in $DataCollection) {
    $TargetDate = [System.DateTime]$Data.Date
    $FieldID = New-Guid
    $Name = "Day" + $TargetDate.ToString("yyyyMMdd")
    $DisplayName = $TargetDate.ToString("M/d (ddd)")
    $Description = ""
    $IsRequired = $True
    $EnforceUniqueValues = $False
    $MaxLength = 255

    $MenuArray = @($Data.Menu1, $Data.Menu2, $Data.Menu3, $Data.Menu4, $Data.Menu5, $Data.Menu6,
        $Data.Menu7, $Data.Menu8, $Data.Menu9, $Data.Menu10, $Data.Menu11, $Data.Menu12, $Data.Menu13)

    $MenuChoiceField = ""
    ForEach ($MenuValuesKey in $MenuValues.Keys) {
        if ([Array]::IndexOf($MenuArray,$MenuValuesKey) -eq -1) {
            continue
        }
        $MenuChoiceField += "<CHOICE>$MenuValuesKey</CHOICE>"

        if($TargetMenuValues.Contains($MenuValuesKey) -eq $False) {
            $TargetMenuValues.Add($MenuValuesKey,$MenuValues[$MenuValuesKey])
        }
    }

    $FieldSchema = @"
        <Field Type='Choice' ID='{$FieldID}' Name='$Name' StaticName='$Name' DisplayName='$DisplayName' Description='$Description' Required='$IsRequired' EnforceUniqueValues='$EnforceUniqueValues' MaxLength='$MaxLength'>
            <CHOICES>
"@ + $MenuChoiceField + @"
            </CHOICES>
        <Default>なし</Default>
        </Field>
"@

    $List.Fields.AddFieldAsXml($FieldSchema, $True, [Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldInternalNameHint)
    $Context.ExecuteQuery()
}

$TargetDate = $InitialDate
$DisplayName = $TargetDate.ToString("M月の合計")

$AmountField = $List.Fields | Where-Object { $_.Title -eq $DisplayName }
if ($Null -ne $AmountField) {
    $AmountField.DeleteObject()
    $Context.ExecuteQuery()
}

$AmountField = $List.Fields | Where-Object { $_.Title -eq $DisplayName }
if ($Null -eq $AmountField) {
    $Fields = $List.Fields
    $Context.Load($Fields)
    $Context.ExecuteQuery()

    $FieldID = New-Guid
    $Name = "Day" + $TargetDate.ToString("yyyyMMdd") + "_Amount"
    $Description = ""
    $ResultType = "Currency"

    $Formula = "="
    ForEach ($Data in $DataCollection) {
        $Count = 0
        $Date = [System.DateTime]$Data.Date
        ForEach ($TargetMenuValueKey in $TargetMenuValues.Keys) {
            $Count++
            $Formula += "IF("
            $Formula += $Date.ToString("[M/d (ddd)]")
            $MenuValue = $TargetMenuValues[($TargetMenuValueKey)]
            $Formula += "=""$TargetMenuValueKey"",$MenuValue"
            if ($Count -eq $TargetMenuValues.Count) {
                $Formula += ",0"
            }
            else {
                $Formula += ","
            }
        }
        For ($MenuCount = 0; $MenuCount -lt $TargetMenuValues.Count; $MenuCount++) {
            $Formula += ")"
        }
        if ($Date -ne $LastDate) {
            $Formula += "+"
        }
    }

    $FieldRefXML=""
    $FieldRefs = @()

    ForEach($Data in $DataCollection) {
        $Date = [System.DateTime]$Data.Date
        $FieldRefs += $Date.ToString("[M/d (ddd)]")
    }

    ForEach($FieldRef in $FieldRefs)
    {
        $FieldRefXML = $FieldRefXML + "<FieldRef Name='$FieldRef' />"
    }

    $Formula.Replace("IF (", "IF(")
    $Formula.Replace("\n", "")
    Log("FormulaLength: " + $Formula.Length)
    $FieldSchema = "<Field Type='Calculated' ID='{$FieldID}' DisplayName='$DisplayName' Name='$Name' Description='$Description' ResultType='$ResultType' ReadOnly='TRUE'><Formula>$Formula</Formula><FieldRefs>$FieldRefXML</FieldRefs></Field>"
    $NewField = $List.Fields.AddFieldAsXml($FieldSchema, $True, [Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldInternalNameHint)
    $Context.ExecuteQuery()

    Log("$DisplayName Field Added")
}
else {
    Log("$DisplayName Field Already Exist")
}

$ADUsers = Get-ADUser -Filter { Enabled -eq $True } -SearchBase "OU=OU, DC=tenant,DC=local" -Property Mail | Where-Object { $_.Surname -ne $Null -and $_.Info -ne "" } | Select-Object Name, Mail

$Query = New-Object Microsoft.SharePoint.Client.CamlQuery
$Query.ViewXml = ""

$ListItems = $List.GetItems($Query)
$Context.Load($ListItems)
$Context.ExecuteQuery()

ForEach ($ADUser in $ADUsers) {
    if($Null -eq $ADUser.Mail) {
        continue;
    }

    $IsExist = $ListItems | Where-Object { $_["Email"] -eq $ADUser.Mail }
    if ($Null -eq $IsExist) {
        $ListItemCreationInformation = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
        $NewListItem = $List.AddItem($ListItemCreationInformation)
        $NewListItem["User"] = $Context.Web.EnsureUser($ADUser.Mail)
        $NewListItem["Email"] = $ADUser.Mail
        ForEach ($Data in $DataCollection) {
            $Date = [System.DateTime]$Data.Date
            $Day = "Day" + $Date.ToString("yyyyMMdd")
            $NewListItem["$Day"] = "なし"
        }
        $NewListItem.Update()
        $Context.ExecuteQuery()
    }
}

$Fields = $List.Fields
$Context.Load($Fields)
$Context.ExecuteQuery()

$NameField = $Fields | Where-Object {$_.InternalName -eq "User"}
$NameField.ReadOnlyField = $True
$NameField.SetShowInEditForm($False)
$NameField.SetShowInNewForm($True)
$NameField.UpdateAndPushChanges($False)
$Context.ExecuteQuery()

$Views = $List.Views
$Context.Load($Views)
$Context.ExecuteQuery()
Log("$List.Views Load Executed")

$TargetYearMonth = $TargetDate.ToString("MM") + "月"

$ViewQuery = @"
<Where>
    <Eq>
        <FieldRef Name="User"/>
        <Value Type="Integer">
        <UserID Type="Integer" />
        </Value>
    </Eq>
</Where>
"@

$ViewCreationInfo = New-Object Microsoft.SharePoint.Client.ViewCreationInformation
$ViewCreationInfo.Title = $TargetYearMonth
$ViewCreationInfo.Query = $ViewQuery
$ViewCreationInfo.RowLimit = "30"

$FieldArray = @("名前")
$FieldArray += $DisplayName
ForEach($Data in $DataCollection) {
    $Date = [System.DateTime]$Data.Date
    $FieldArray += $Date.ToString("M/d (ddd)")
}

$ViewCreationInfo.ViewFields = $FieldArray
$ViewCreationInfo.SetAsDefaultView = $True

$NewView = $Views.Add($ViewCreationInfo)
$Context.ExecuteQuery()
Log("$DisplayName View Added")

$Context.Dispose()
Disconnect-SPOService
Log("Disconnect-SPOService")
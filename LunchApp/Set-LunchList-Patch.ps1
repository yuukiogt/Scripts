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

$SharePointModule = Get-Module -Name Microsoft.Online.SharePoint.PowerShell -ListAvailable | Select-Object Name, Version
if ($Null -eq $SharePointModule) {
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser
    Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Force -Scope CurrentUser
}

$UserName = $env:UserName + "@tenant.co.jp"
$M365Credential = Get-Credential -UserName $UserName -Message "Enter Password"
Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
Connect-SPOService -Url https://tenant-admin.sharepoint.com -Credential $M365Credential | Out-Null

[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime") | Out-Null

$SiteUrl = "https://tenant.sharepoint.com/sites/site/"
$Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
$SPCredential = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($M365Credential.UserName, $M365Credential.Password)
$Context.Credentials = $SPCredential
$WebObject = $Context.Web
$Context.Load($WebObject)
$Context.ExecuteQuery()
Log("ClientContext Executed")

$ListName = "お弁当注文_" + (Get-Date).AddMonths(1).ToString("yyyy")
$List = $Context.Web.Lists.GetByTitle($ListName)
$Context.Load($List)
$Context.ExecuteQuery()
Log("$ListName GetByTitle Executed")

$IsEnableDate = $False
$TargetDate = $Null

do {
    $TargetDateString = Read-Host "修正する日を入力してください (例:7/1)"
    $TargetDate = [System.DateTime]$TargetDateString

    if($Null -ne $TargetDate) {
        $IsEnableDate = $True
    }
} while ($IsEnableDate -eq $False)

$Fields = $List.Fields
$Context.Load($Fields)
$Context.ExecuteQuery()

$Field = $List.Fields | Where-Object { $_.Title -eq ($TargetDate.ToString("M/d (ddd)")) }
$Field.ReadOnlyField = $False
$Field.Update()
$Context.ExecuteQuery()

$Wait = Read-Host "$TargetDate.ToString(""M/d"") の列が修正可能になりました。修正が完了したら Enterを押してください"

$Field.ReadOnlyField = $True
$Field.Update()
$Context.ExecuteQuery()

$Context.Dispose()
Disconnect-SPOService
Log("Disconnect-SPOService")
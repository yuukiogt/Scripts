$ErrorActionPreference = "SilentlyContinue"

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
if($Null -eq $SharePointModule) {
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser
    Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Force -Scope CurrentUser
}

$UserName = $env:UserName + "@tenant.co.jp"
$M365Credential = Get-Credential -UserName $UserName -Message "Enter Password"
Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking

[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime") | Out-Null

$SiteUrl = "https://tenant.sharepoint.com/sites/site/"
$Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
$SPCredential = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($M365Credential.UserName, $M365Credential.Password)
$Context.Credentials = $SPCredential
$WebObject = $Context.Web
$Context.Load($WebObject)
$Context.ExecuteQuery()

Function Invoke-LoadMethod() {
    param(
        [Microsoft.SharePoint.Client.ClientObject]$Object = $(throw "Please provide a Client Object"),
        [string]$PropertyName
    ) 
    $Context = $Object.Context
    $Load = [Microsoft.SharePoint.Client.ClientContext].GetMethod("Load") 
    $Type = $Object.GetType()
    $ClientLoad = $Load.MakeGenericMethod($Type)

    $Parameter = [System.Linq.Expressions.Expression]::Parameter(($Type), $Type.Name)
    $Expression = [System.Linq.Expressions.Expression]::Lambda([System.Linq.Expressions.Expression]::Convert([System.Linq.Expressions.Expression]::PropertyOrField($Parameter, $PropertyName), [System.Object] ), $($Parameter))
    $ExpressionArray = [System.Array]::CreateInstance($Expression.GetType(), 1)
    $ExpressionArray.SetValue($Expression, 0)
    $ClientLoad.Invoke($Context, @($Object, $ExpressionArray))
}

$ListName = "お弁当注文_" + (Get-Date).AddMonths(1).ToString("yyyy")
$List = $Context.Web.Lists.GetByTitle($ListName)
$Context.Load($List)
$Context.ExecuteQuery()

$TargetMonth = $Null

do {
    [int]$TargetDateString = Read-Host "集計ビューの月を入力してください 1～12"
    if ($TargetDateString -ge 1 -and $TargetDateString -le 12) {
        $TargetMonth = Get-Date -Month $TargetDateString -Day 1
    }
    if ($TargetDateString -eq 1) {
        $TargetMonth = $TargetMonth.AddMonths(12)
    }
} while ($Null -eq $TargetMonth)

$Views = $List.Views
$Context.Load($Views)
$Context.ExecuteQuery()

$IsExist = $List.Views.GetByTitle(($TargetMonth.ToString("集計_MM月")))
$OverWrite = $False
if($Null -ne $IsExist) {
    $OverWrite = Read-Host "既に存在しています。上書きしますか？ (y/n)"
    if($OverWrite -ne "y") {
        Exit
    }
}

if($OverWrite -eq "y") {
    $View = $List.Views.GetByTitle(($TargetMonth.ToString("集計_MM月")))
    $Context.Load($View)
    $Context.ExecuteQuery()

    $View.DeleteObject()
    $Context.ExecuteQuery()
}

$View = $List.Views.GetByTitle(($TargetMonth.ToString("MM月")))
$Context.Load($View)
$Context.ExecuteQuery()

Invoke-LoadMethod -Object $View -PropertyName "ViewFields"
$Context.ExecuteQuery()

$ViewCreationInfo = New-Object Microsoft.SharePoint.Client.ViewCreationInformation
$ViewCreationInfo.ViewFields = $View.ViewFields
$ViewCreationInfo.Title = "集計_" + $View.Title
$ViewCreationInfo.Query = ""
$ViewCreationInfo.RowLimit = "300"
$ViewCreationInfo.PersonalView = $True
$ViewCreationInfo.SetAsDefaultView = $False
$ViewCreationInfo.Paged = $View.Paged
$ViewCreationInfo.ViewTypeKind = $View.ViewType

$Views.Add($ViewCreationInfo) | Out-Null
$Context.ExecuteQuery()

$Context.Dispose()

Log("集計ビューを追加しました")
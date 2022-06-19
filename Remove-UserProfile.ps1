# https://social.technet.microsoft.com/Forums/ja-JP/af03f04f-2e74-4717-8d8c-88851938119a/rd312272120512518125401247012540125031252512501124491245212523?forum=activedirectoryja

function RemoveUserProfile() {
    Param(
    [Parameter(Mandatory=$True,Position=1)]
    [string]$Name,
    [Parameter(Mandatory=$False,Position=2)]
    [string]$ComputerName="."
    )

    $objDomain = New-Object System.DirectoryServices.DirectoryEntry
    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
    $objSearcher.SearchRoot = $objDomain
    $objSearcher.SearchScope = "Subtree"

    $strFilter = "(cn=$Name)"
    $objSearcher.Filter = $strfilter
    $objUser = ($objSearcher.Findone()).GetDirectoryEntry()
    $SID=(New-Object System.Security.Principal.SecurityIdentifier $objUser.objectsid[0],0).Value

    Invoke-Command -Computername $ComputerName {
        Param($Name,$SID)
        Set-Location "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
        IF ((Test-Path $SID) -eq $True) {
            $ProfilePath=(Get-Item $SID).GetValue("ProfileImagePath")
            $RegistryPath=(Get-Item $SID).PSPath
            cmd /C rmdir /S /Q $ProfilePath
            Remove-Item $RegistryPath -Recurse -Force
            Write-Host $Name " (" $SID ") のプロファイルを削除しました。"
            }
        Else	{
            Write-Host $Name " (" $SID ") のプロファイルは見つけられませんでした。"
            }
    } -ArgumentList $Name,$SID
}

$userName = Read-Host "ユーザー名"
$computerName = Read-Host "コンピューター名"

RemoveUserProfile -Name $userName -ComputerName $computerName
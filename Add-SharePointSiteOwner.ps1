$ErrorActionPreference = "Inquire"

$CurrentDir = Split-Path $MyInvocation.MyCommand.Path -Parent
$PasswordFile = Join-Path $CurrentDir "administrator.txt"

Get-Credential
$UserName = "administrator@tenant.onmicrosoft.com"
$SecurePassword = Get-Content $PasswordFile | ConvertTo-SecureString -Key (1..16)

$M365Credential = New-Object System.Management.Automation.PSCredential $UserName, $SecurePassword
Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
Connect-SPOService -Url https://tenant-admin.sharepoint.com -Credential $M365Credential | Out-Null
Connect-AzureAD -Credential $M365Credential | Out-Null

$Sites = Get-SPOSite -Limit All

[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime") | Out-Null

$admin = Get-AzureADUser -Filter "userPrincipalName eq '$($UserName)'"

$add = $False

if ((Read-Host "add? (y:add /n: remove)") -eq 'y') {
    $add = $True
}

ForEach ($Site in $Sites) {
    $SiteURL = $Site.Url
    Write-Host -f Cyan $SiteURL

    Try {
        $GroupOwners = Get-AzureADGroupOwner -ObjectId $Site.GroupID | Select-Object -ExpandProperty UserPrincipalName
        if(!$GroupOwners.Contains($UserName) -and $add) {
            Add-AzureADGroupOwner -ObjectId $Site.GroupID -RefObjectId $admin.ObjectId
        } else {
            Remove-AzureADGroupOwner -ObjectId $Site.GroupID -OwnerId $admin.ObjectId
        }

        $Groups = Get-SPOSiteGroup -Site $SiteURL |
        Where-Object { !$_.Users.Contains($UserName) -and !$_.Title.StartsWith("SharingLinks") -and !$_.Title.StartsWith("Limited Access System Group") -and !$_.Title.Contains("SHAREPOINT\System") }

        ForEach($Group in $Groups)
        {
            Write-Host -f Yellow $Group.Title
            Write-Host -f Yellow $Group.Users

            if(!$Group.Users.Contains($UserName) -and $add) {
                Add-SPOUser -Site $SiteURL -LoginName $UserName -Group $Group.Title
            } else {
                Remove-SPOUser -Site $SiteURL -LoginName $UserName -Group $Group.Title
            }
        }
    }
    Catch {
        Write-Host -f Red $_.Exception.Message
    }
    Finally {
    }
}

Disconnect-AzureAD
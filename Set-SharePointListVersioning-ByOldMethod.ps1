$ErrorActionPreference = "Inquire"

$M365Credential = Get-Credential
Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
Connect-SPOService -Url https://tenant-admin.sharepoint.com -Credential $M365Credential | Out-Null
Connect-AzureAD -Credential $M365Credential | Out-Null

$Sites = Get-SPOSite -Limit All

[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime") | Out-Null

$VersionsToKeep = 10

ForEach ($Site in $Sites) {
    $SiteURL = $Site.Url

    Try {
        $Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
        $Ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName, $SecurePassword)

        $Lists = $Ctx.Web.Lists
        $Ctx.Load($Lists)
        $Ctx.ExecuteQuery()

        Write-Host -f Yellow "Processing Site: "$SiteURL

        $Lists = $Lists | Where-Object { $_.BaseType -eq "DocumentLibrary" -and $_.EnableVersioning -and $_.Hidden -eq $False }

        $Query = New-Object Microsoft.SharePoint.Client.CamlQuery
        $Query.ViewXml = "<View Scope='RecursiveAll'><Query><Where><Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>0</Value></Eq></Where></Query></View>"

        ForEach ($List in $Lists) {            
            $ListItems = $List.GetItems($Query)
            $Ctx.Load($ListItems)
            $Ctx.ExecuteQuery()
    
            Foreach ($Item in $ListItems) {
                $Ctx.Load($Item)
                $Ctx.ExecuteQuery()

                $Versions = $Item.File.Versions
                $Ctx.Load($Versions)
                $Ctx.Load($Item.File)
                $Ctx.ExecuteQuery()
        
                Write-host -f Yellow "Total Number of Versions Found in '$($Item.File.Name )' : $($Versions.count)"
        
                While ($Item.File.Versions.Count -gt $VersionsToKeep) {
                    write-host "Deleting Version:" $Versions[0].VersionLabel
                    $Versions[0].DeleteObject()
                    $Ctx.ExecuteQuery()
            
                    $Ctx.Load($Item.File.Versions)
                    $Ctx.ExecuteQuery()
                }
            }

            $List.EnableVersioning = $true
            $List.EnableMinorVersions = $false
            $List.MajorVersionLimit = 10
            $List.Update()
        }
    }
    Catch {
        Write-Host -f Red $_.Exception.Message
        Read-Host "press any key..."
    }
    Finally {
    }
}
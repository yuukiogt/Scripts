try {
    $M365Credential = Get-Credential
    Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
    Connect-SPOService -Url https://tenant-admin.sharepoint.com -Credential $M365Credential | Out-Null

    $Sites = Get-SPOSite -Limit All

    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client") | Out-Null
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime") | Out-Null

    $SiteUrl = "https://tenant.sharepoint.com/sites/nic/"
    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
    $SPCredential = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName, $SecurePassword)
    $Context.Credentials = $SPCredential
    $Context.ExecuteQuery()

    $ListName = "SharePoint 使用状況"
    $List = $Context.Web.Lists.GetByTitle($ListName)
    $Context.Load($List)
    $Context.ExecuteQuery()

    $Query = New-Object Microsoft.SharePoint.Client.CamlQuery
    $Query.ViewXml = ""

    do {
        $ListItems = $List.getItems($Query)
        $Context.Load($ListItems)
        $Context.ExecuteQuery()

        while ($ListItems.Count -gt 0) {
            $ListItems[0].DeleteObject()
        }

        $Query.ListItemCollectionPosition = $items.ListItemCollectionPosition
    }
    while ($null -ne $Query.ListItemCollectionPosition)
    $Context.ExecuteQuery()

    $M365Credential = Get-Credential
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $M365Credential -Authentication Basic -AllowRedirection
    Import-PSSession $Session

    $Groups = Get-UnifiedGroup -ResultSize Unlimited

    foreach ($Site in $Sites) {

        if ($Site.Title -eq "") {
            continue;
        }

        if (!$Site.Url.Contains("https://tenant.sharepoint.com/sites")) {
            continue;
        }

        if ($Site.Title.Contains("RedirectSite")) {
            continue;
        }

        if ($Site.Status -ne "Active") {
            continue;
        }

        $Owner = ""
        $TargetGroup = $null
        foreach ($Group in $Groups) {
            if ($Site.Title -eq $Group.DisplayName -or $Site.Url -eq $Group.SharePointSiteUrl) {
                $Owners = $Group | Get-UnifiedGroupLinks -LinkType Owners
                $TargetGroup = $Group
                if ($null -eq $Owners) {
                    $Owner = "所有者なし"
                    break
                }
                foreach ($o in $Owners) {
                    if($o.DisplayName -eq "全体管理者") {
                        continue
                    }
                    $Owner += $o.DisplayName + ", "
                }
                $Owner = $Owner.Remove($Owner.Length - 2, 2)
                break
            }
        }

        if ($Owner -eq "") {
            $Owner = (Get-User -Identity $Site.Owner).DisplayName
        }
        
        $ListItemCreationInformation = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
        $NewListItem = $List.AddItem($ListItemCreationInformation)
        $NewListItem["Title"] = $Site.Title
        $NewListItem["StorageUsageCurrent"] = $Site.StorageUsageCurrent / 1024
        $NewListItem["Url"] = $Site.Url
        $NewListItem["LastContentModifiedDate"] = $TargetGroup.WhenChanged
        $NewListItem["Owner"] = $Owner
        $NewListItem.Update()
        $Context.ExecuteQuery()
    }

} catch {
    Write-Host $_.Exception.Message
} finally {
    if($Session) {
        Remove-PSSession $Session
    }
    if($Context) {
        $Context.Dispose()
    }

    Disconnect-SPOService
}
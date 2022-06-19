 # Read CSOM
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime") | Out-Null

# SiteCollection URL 
$SiteCollectionUrl = "https://tenant.sharepoint.com/sites/site/" 

# Account
$Account = Read-Host -Prompt "Enter Your Account."
$SecurePassword = Read-Host -Prompt "Enter Your Password." -AsSecureString

# Create Credential
$Credential = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Account, $SecurePassword) 

# Create Context
$Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteCollectionUrl)
$Context.Credentials = $credential

$objListName = "Document Library"

# Create List object
$objList = $Context.Web.Lists.GetByTitle($objListName)
$Context.Load($objList)
$Context.ExecuteQuery()

Write-Host "list Title : " $objList.Title
Write-Host "list BaseType : " $objList.BaseType
Write-Host "list Created : " $objList.Created

$query = New-Object Microsoft.SharePoint.Client.CamlQuery
$query.ViewXml = ""

$objListItems = $objList.getItems($query)
$Context.Load($objListItems)
$Context.ExecuteQuery()

Write-Host ""

$totalSize = 0

foreach($item in $objListItems)
{
    $totalSize += ($item.FieldValues.File_x0020_Size / 1KB)
}

Write-Host $totalSize KB

$Context.Dispose()


$ErrorActionPreference = "SilentlyContinue"

$M365Credential = Get-Credential
Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
Connect-SPOService -Url https://tenant-admin.sharepoint.com -Credential $M365Credential | Out-Null

$Sites = Get-SPOSite -Limit All |
Select-Object LastContentModifiedDate, Status, Title, Url |
Where-Object { $_.Status -eq "Active" -and $_.Title -ne "" }

[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime") | Out-Null

$VersionsToKeep = 10
$ListName = "ドキュメント"

function Add-UserAgent($ctx) {
  $clientAssembly = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
  $clientRuntimeAssembly = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
  $assemblies = ($clientAssembly.FullName, $clientRuntimeAssembly.FullName)
  Add-Type -Language CSharp -ReferencedAssemblies $assemblies -TypeDefinition @"
    using System; 
    using Microsoft.SharePoint.Client;

    public static class SPContextHelper
    {
        public static void AddUserAgent(ClientContext ctx)
        {
            ctx.ExecutingWebRequest += delegate (object sender, WebRequestEventArgs e)
            {
                e.WebRequestExecutor.WebRequest.UserAgent = "NONISV|company|DummyScript/1.0";
            };
        }
    }
"@
  [SPContextHelper]::AddUserAgent($ctx);
}

function ExecuteQueryWithIncrementalRetry {
  param (
    [parameter(Mandatory = $true)]
    [int]$retryCount,
    $ctx
  );

  $DefaultRetryAfterInMs = 120000;
  $RetryAfterHeaderName = "Retry-After";
  $retryAttempts = 0;

  if ($retryCount -le 0) {
    throw "Provide a retry count greater than zero."
  }

  while ($retryAttempts -lt $retryCount) {
    try {
      $ctx.ExecuteQuery();
      return;
    }
    catch [System.Net.WebException] {
      $response = $_.Exception.Response

      if (($null -ne $response) -and (($response.StatusCode -eq 429) -or ($response.StatusCode -eq 503))) {
        $retryAfterHeader = $response.GetResponseHeader($RetryAfterHeaderName);
        $retryAfterInMs = $DefaultRetryAfterInMs;

        if (-not [string]::IsNullOrEmpty($retryAfterHeader)) {
          if (-not [int]::TryParse($retryAfterHeader, [ref]$retryAfterInMs)) {
            $retryAfterInMs = $DefaultRetryAfterInMs;
          }
          else {
            $retryAfterInMs *= 1000;
          }
        }

        Write-Output ("CSOM request exceeded usage limits. Sleeping for {0} seconds before retrying." -F ($retryAfterInMs / 1000))
        Start-Sleep -m $retryAfterInMs
        $retryAttempts++;
      }
      else {
        throw;
      }
    }
  }

  throw "Maximum retry attempts {0}, have been attempted." -F $retryCount;
}

ForEach ($Site in $Sites) {
    $SiteURL = $Site.Url

    $SiteURL

    Try {
        Connect-PnPOnline -Url $SiteURL -Credentials $M365Credential
        $Ctx = Get-PnPContext
        
        Add-UserAgent $Ctx

        $ListItems = Get-PnPListItem -List $ListName -PageSize 2000 | Where-Object { $_.FileSystemObjectType -eq "File" }

        ForEach ($Item in $ListItems) {
            $File = $Item.File
            $Versions = $File.Versions
            $Ctx.Load($File)
            $Ctx.Load($Versions)
            ExecuteQueryWithIncrementalRetry -ctx $Ctx -retryCount 10

            Write-host -f Yellow "Scanning File:"$File.Name
            $VersionsCount = $Versions.Count
            $VersionsToDelete = $VersionsCount - $VersionsToKeep

            If ($VersionsToDelete -gt 0) {
                write-host -f Cyan "`t Total Number of Versions of the File:" $VersionsCount
                For ($i = 0; $i -lt $VersionsToDelete; $i++) {
                    write-host -f Cyan "`t Deleting Version:" $Versions[0].VersionLabel
                    $Versions[0].DeleteObject()
                }
                try {
                    ExecuteQueryWithIncrementalRetry -ctx $Ctx -retryCount 10
                }
                Catch {
                    Write-Host -f Red $_.Exception.Message
                }
                Write-Host -f Green "`t Version History is cleaned for the File:"$File.Name
            }
        }

        $Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($site.Url)
        $Ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName, $securePassword)

        $Lists = $Ctx.Web.Lists
        $Ctx.Load($Lists)
        ExecuteQueryWithIncrementalRetry -ctx $Ctx -retryCount 10

        $Lists = $Lists | Where-Object { $_.BaseType -eq "DocumentLibrary" -and $_.EnableVersioning -and $_.Hidden -eq $False }

        $Query = New-Object Microsoft.SharePoint.Client.CamlQuery
        $Query.ViewXml = "<View Scope='RecursiveAll'><Query><Where><Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>0</Value></Eq></Where></Query></View>"

        ForEach ($List in $Lists) {
          $List.EnableVersioning = $true
          $List.EnableMinorVersions = $false
          $List.MajorVersionLimit = 10
          $List.Update()
        }
    }
    Catch {
        Write-Host -f Red $_.Exception.Message
    }
}
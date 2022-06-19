$ErrorActionPreference = "SilentlyContinue"

try {
    $connection = Connect-ExchangeOnline -Credential (Get-Credential)

    $csvPath = ".\Users.csv"

    $dl = "distributionList"

    Import-Csv $csvPath | Foreach-Object {
        New-ComplianceSearch -Name "DidArrivedEmailTo_$($dl)" -ExchangeLocation $dl -ContentMatchQuery "subject:$($subject)"
    }
} catch {
    $_.Exception.Message
} finally {
    if($Null -ne $connection) {
        Disconnect-ExchangeOnline
    }
}
$Credential = Get-Credential

$OU = "*"

try {
    Import-Module AzureAD -Force
    $connect = Connect-AzureAD -Credential $Credential

    $azDevices = @()
    Get-AzureADDevice -All:$True | Where-Object {$_.AccountEnabled -eq $True} | Select-Object DisplayName |
    Foreach-Object {
        $azDevices += $_.DisplayName
    }

    $azDevicesDic = @{}
    Get-AzureADDevice -All:$True | Where-Object { $_.AccountEnabled -eq $True } | Select-Object ObjectId, DisplayName |
    Foreach-Object {
        $azDevicesDic[$_.DisplayName] = $_.ObjectId
    }

    $adDevices = @()
    Get-ADComputer -Filter { Enabled -eq $True } | Where-Object { $_.DistinguishedName -like $OU } | Select-Object Name |
    ForEach-Object {
        $adDevices += $_.Name
    }

    $diffs = $azDevices | Where-Object { $adDevices -notcontains $_ }
    if($diffs.Length -eq 0) {
        Write-Host "AzureAD上にオンプレADで登録されていないコンピューターはありません"
    } else {
        Write-Host "diffs: $($diffs)"
    }
    
    foreach ($diff in $diffs) {
        if($azDevicesDic.Contains($diff)) {
            do {
                $answer = Read-Host "$diff を無効にしますか? (y/n)"
                if($answer -eq 'y') {
                    Set-AzureADDevice -ObjectId $azDevicesDic[$diff] -AccountEnabled $False
                } elseif($answer -eq 'n') {
                    continue
                } else {
                    # do nothing
                }
            }
            until($answer -eq 'y' -or $answer -eq 'n')
        }
    }
    
} catch {
    $_.Exception.Message
} finally {
    if($Null -ne $connect) {
        Disconnect-AzureAD
    }
}
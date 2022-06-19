$isExistWiFi = Get-NetAdapter | Where-Object { $_.Name -eq "Wi-Fi"}

if($Null -eq $isExistWiFi) {
    Write-Host "Wi-Fiが存在しません。"
} else {
    $upAdapters = Get-NetAdapter | Where-Object { $_.Status -eq "up"}
    foreach($upAdapter in $upAdapters) {
        if($upAdapter.Name -eq "Wi-Fi") {
            continue;
        }
        Disable-NetAdapter $upAdapter.Name -Confirm:$False
    }

    netsh wlan show profile

    netsh wlan connect name="guest"

    Start-Sleep 3
    Write-Host "VPNに接続前"

    $uri = "http://innersite.tenant.co.jp"
    (Invoke-WebRequest -Method Get -Uri $uri).StatusCode
    if($statusCode -eq 200) {
        Write-Host "OK"
    }
    else {
        Write-Host "NG"
    }

    $VpnName = "vpnName"

    $RasExec = "C:\windows\system32\rasdial.exe"

    $VpnUser = ""
    $VpnPass = ""

    do {
        $isValid = $True
        $User = Read-Host @"
            vpn-01: 1
            vpn-02: 2
            vpn-03: 3

"@

        if($User -eq 1) {
            $VpnUser = "vpn-01"
            $VpnPass = "vpn01pass"
        }
        elseif ($User -eq 2) {
            $VpnUser = "vpn-02"
            $VpnPass = "vpn02pass"
        }
        elseif ($User -eq 3) {
            $VpnUser = "vpn-03"
            $VpnPass = "vpn03pass"
        }
        else {
            $isValid = $False
        }
    }
    while($isValid -eq $False)

    cmd.exe /c $RasExec $VpnName $VpnUser $VpnPass
    Write-Host "VPNに接続後"

    netsh wlan connect name="guest"

    Start-Sleep 3

    $uri = "http://innersite.tenant.co.jp"
    (Invoke-WebRequest -Method Get -Uri $uri).StatusCode
    if($statusCode -eq 200) {
        Write-Host "OK"
    }
    else {
        Write-Host "NG"
    }

    foreach($upAdapter in $upAdapters) {
        if($upAdapter.Name -eq "Wi-Fi") {
            continue;
        }
        Enable-NetAdapter $upAdapter.Name
    }

    netsh wlan disconnect interface="Wi-Fi"

    $connectName = Read-Host "無線の接続名？"
    netsh wlan connect name=$connectName

    cmd.exe /c $RasExec /Disconnect
}
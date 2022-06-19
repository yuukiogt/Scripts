function GetOSVersion() {
    $OS = Get-WmiObject Win32_OperatingSystem
    $buildNumber = $OS.BuildNumber
    $versionString = $OS.Version.Replace( ".$buildNumber", "" )
    $version = [decimal]$versionString
    if ( $version -lt 6.0 ) {
        # not support version
        return 0
    }
    elseif (($version -ge 6.0) -and ($version -lt 6.1)) {
        # Windows Vista or Windows Server 2008
        return 1
    }
    elseif (($version -ge 6.1) -and ($version -lt 6.2)) {
        # Windows 7 or Windows Server 2008 R2
        return 7
    }
    elseif (($version -ge 6.2) -and ($version -lt 6.3)) {
        # Windows 8 or Windows Server 2012
        return 8
    }
    elseif (($version -ge 6.3) -and ($version -lt 6.4)) {
        # Windows 8.1 or Windows Server 2012 R2
        return 8.1
    }
    else {
        # Windows 10 and later or Windows Server 2016 and later
        return 10
    }
}

$OSVersion = GetOSVersion

if ($OSVersion -eq 10) {
    $ifIndexes = Get-NetAdapter | Select-Object ifIndex
    foreach ($ifIndex in $ifIndexes) {
        $index = $ifIndex.ifIndex 
        Set-DnsClientServerAddress -InterfaceIndex $index -ServerAddresses "192.168.33.3"
    }
}
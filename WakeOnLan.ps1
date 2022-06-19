$mac_addr = @("")
$header = [byte[]](@(0xFF) * 6)

foreach ($item in $mac_addr) {
    Write-Host "send magic packet to:" $item
    $addr = [byte[]]($item.split(":") | ForEach-Object { [Convert]::ToInt32($_, 16) });
    $magicpacket = $header + $addr * 16;
    $target = [System.Net.IPAddress]::Broadcast;

    $client = New-Object System.Net.Sockets.UdpClient;
    $client.Connect($target, 2304);

    $client.Send($magicpacket, $magicpacket.Length) | Out-Null
    $client.Close();

    Write-Host "Send magic packet to:" $item -ForegroundColor Green
}
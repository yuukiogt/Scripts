$resetTime = 600
$restartSpan = 180000

while(1) {
    $computerName = Read-Host "computerName or q(quit)"
    if ($computerName -eq 'q') {
        exit
    }

    if (-not (Test-Connection $computerName -Count 1)) {
        continue
    }
    
    while(1) {
        $serviceName = Read-Host "serviceName or q(quit)"
        if ($serviceName -eq 'q') {
            break
        }

        $service = Get-Service $serviceName
        if ($Null -eq $service) {
            continue
        }

        $target = "\\$($computerName)"

        sc.exe $target failure $serviceName reset=$resetTime actions=restart/$restartSpan
    }
}
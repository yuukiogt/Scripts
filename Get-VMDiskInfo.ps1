$server = ''
$parent = Split-Path $MyInvocation.MyCommand.Path -Parent
$now = Get-Date -Format 'yyyyMMdd_HHmmss'
$outputPath = Join-Path $parent "VMDiskInfo_$($now).csv"

try {
    Connect-VIServer -Server $server -User 'user@vsphere.local' -Password '' -Force

    Add-Content -Path $outputPath -Value '"HostName","IPv4Address","ProvisionedSpaceGB(GB)","UsedSpace(GB)","Drive","Capacity(GB)","FreeSpace(GB)","Used(GB)","Usage(%)"'

    $VMs = Get-VM

    foreach($vm in $VMs){
        $disks = $vm.Guest.Disks

        foreach($disk in $disks){
            $usage = (100 - ([float]$disk.FreeSpaceGB / [float]$disk.CapacityGB) * 100).ToString("0.0")
            $used = ($disk.CapacityGB - $disk.FreeSpaceGB).ToString("0.0")
            Add-Content -Path $outputPath -Value "$($vm.name),$($vm.Guest.IPAddress[0]),$($vm.ProvisionedSpaceGB.ToString("0.0")),$($vm.UsedSpaceGB.ToString("0.0")),$($disk.path),$($disk.CapacityGB.ToString("0.0")),$($disk.FreeSpaceGB.ToString("0.0")),$($used),$($usage)"
        }
    }
} catch {
    Write-Host $_.Exception.Message
} finally {
    Disconnect-VIServer -Server $server -Confirm:$false
}
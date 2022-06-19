$log = ".\$($env:ComputerName)_$(Get-date -f yyyyMMdd).txt"
Start-Transcript $log -Append

Write-Host "適用KB一覧"
Get-Hotfix | Format-Table HotFixID, InstalledOn -AutoSize

$updateSession = New-Object -com Microsoft.Update.Session

Write-Host "Windows Update を検索しています..."
$searcher = $updateSession.CreateUpdateSearcher()
$searchResult = $searcher.search("IsInstalled=0 and Type='software'")

$result = $searchResult.Updates | ForEach-Object { $_.title -replace ".*(KB\d+).*", "`$1`t$&" }
$result
if (($result | Measure-Object).Count -eq 0) {
    Write-Host "適用が必要な Update はありません。"
    Write-Host "適用KB一覧"
    Get-Hotfix | Format-Table HotFixID, InstalledOn -AutoSize
    exit 0
}

$updatesToDownload = New-Object -com Microsoft.Update.UpdateColl
$searchResult.Updates | Where-Object { -not $_.InstallationBehavior.CanRequestUserInput } |
Where-Object { $_.EulaAccepted } |
ForEach-Object { [void]$updatesToDownload.add($_) }
$updatesToDownload | Format-List Title

Write-Host "ダウンロードしています..."
$downloader = $updateSession.CreateUpdateDownloader()
$downloader.Updates = $updatesToDownload
$downloader.Download()

$updatesToInstall = New-Object -com Microsoft.Update.UpdateColl
$searchResult.Updates | Where-Object { $_.IsDownloaded } | ForEach-Object { [void]$updatesToInstall.add($_) }
$updatesToInstall | Format-List Title

Write-Host "インストールしています..."
$installer = $updateSession.CreateUpdateInstaller()
$installer.Updates = $updatesToInstall
$installationResult = $installer.Install()

Write-Host "結果:（2 なら成功、3以上なら一部または全て失敗）"
$installationResult
if ($installationResult.RebootRequired -eq $true) {
    Write-Host "再起動が必要です"
    Get-date
    Stop-Transcript
    exit 1
}
else {
    Stop-Transcript
    exit 2
}
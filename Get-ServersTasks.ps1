$servers = Get-ADComputer -Filter { OperatingSystem -like "*Server*" -and Enabled -eq $True }

$dest = [Environment]::GetFolderPath("Desktop")
$csvPath = Join-Path $dest "ServerTaskScheduler.csv"
$isExistClientsCSV = Test-Path $csvPath

if ($isExistClientsCSV -eq $False) {
    Add-Content -Path $csvPath -Encoding UTF8 -Value '"ホスト名","タスク名","次回の実行時刻","状態","ログオンモード","前回の実行時刻","前回の結果","作成者","実行するタスク","開始","コメント","スケジュールされたタスクの状態","アイドル時間","電源管理","ユーザーとして実行","再度スケジュールされない場合はタスクを削除する","タスクを停止するまでの時間","スケジュール","スケジュールの種類","開始時刻","開始日","終了日","日","月","繰り返し: 間隔","繰り返し: 終了時刻","繰り返し: 期間","繰り返し: 実行中の場合は停止"'
}

$csv = Import-Csv -Encoding UTF8 -Path $csvPath

$creator = ""

foreach ($server in $servers) {
    Write-Host "${server} のタスクを出力しています ..."
    $content = schtasks.exe /query /v /s $server.Name /FO csv | findstr $creator
    Add-Content -Path $csvPath -Encoding UTF8 -Value $content
    Write-Host "${server} のタスクを出力しました"
}
$srcPath = ".\Src"
$dstPath = ".\ScanResults"

$targetFiles = Join-Path $srcPath "*_emocheck.txt"
Move-Item -Path $targetFiles -Destination $dstPath -Force

$files = Get-ChildItem -Path $dstPath -File

foreach ($file in $files) {
    foreach ($str in Get-Content $file.FullName -Encoding UTF8 -Tail 1) {
        if ($str.Contains("検知しませんでした") -or $str.Contains("Emotet was not detected.")) {
            $file | Rename-Item -NewName { $_.Name -replace '_emocheck.txt', '_emocheck_OK_.txt' }
        }
        else {
            $file | Rename-Item -NewName { $_.Name -replace '_emocheck.txt', '_emocheck_Detected_.txt' }            
        }
    }
}
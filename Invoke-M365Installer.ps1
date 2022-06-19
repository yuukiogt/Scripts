try {
    do {
        $Target = Read-Host @"
        
        セットアップする対象を選択してください。

        0: all
        1: 横浜 64bit

        q: 終了

"@
        switch ($Target) {
            0 {

            }
            1 { 
                .\setup.exe /download Configuration-Monthly.xml
            }
            Default {}
        }
    } while( $Target -ne "q")

} catch {
    $_.Exception.Message
}
function SetMouseSpeed() {
    $newSpeed = 20

    $winApi = Add-Type -Name user32 -Namespace tq84 -PassThru -MemberDefinition '
    [DllImport("user32.dll")]
    public static extern bool SystemParametersInfo(
       uint uiAction,
       uint uiParam ,
       uint pvParam ,
       uint fWinIni
    );
'

    $SPI_SETMOUSESPEED = 0x0071
    $null = $winApi::SystemParametersInfo($SPI_SETMOUSESPEED, 0, 20, 0)

    Set-ItemProperty 'HKCU:\Control Panel\Mouse' -Name MouseSensitivity -Value $newSpeed
}

function SetWindowsSettings() {
    #拡張子を表示する
    Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name "HideFileExt" -Value 0
    
    # 視覚効果 設定プロファイル カスタム
    Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects' -Name "VisualFXSetting" -Value 3
    
    # Windows内のアニメーションコントロールと要素
    # ウィンドウの下に影を表示する
    # コンボボックスをスライドして開く
    # ヒントをフェードまたはスライドで表示する
    # マウスポインターの下に影を表示する
    # メニューをフェードまたはスライドして表示する
    # メニュー項目をクリック後にフェードアウトする
    # リストボックスを滑らかにスクロールする
    Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name "UserPreferencesMask" -Value ([byte[]](0x00, 0x00, 0x00, 0x00, 0x10, 0x00, 0x00, 0x00))

    # ウィンドウを最大化や最小化するときにアニメーションで表示する
    Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop\WindowMetrics' -Name "MinAnimate" -Value 0
    # アイコンの代わりに縮小版を表示する
    Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name "IconsOnly" -Value 0
    # タスクバーとスタートメニューでアニメーションを表示する
    Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name "TaskbarAnimations" -Value 0
    # デスクトップのアイコン名に影を付ける
    Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name "ListviewShadow" -Value 0
    # 半透明の［選択］ツールを表示する
    Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name "ListviewAlphaSelect" -Value 0
    # ウィンドウとボタンに視覚スタイルを使用する
    Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\ThemeManager' -Name "ThemeActive" -Value 1
 
    # Windows の表示に透明性を適用する をオフにする
    Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -Name "EnableTransparency" -Value 0

    # 隠しファイルを表示する
    Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name "Hidden" -Value 1
    Stop-Process -processname explorer
}

SetMouseSpeed
SetWindowsSettings
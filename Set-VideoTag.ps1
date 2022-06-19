try {
    Import-Module -Name Selenium -ErrorAction Stop

    $driver = Start-SeChrome

    if (!$driver) {
        Write-Host "The selenium driver was not running." -ForegroundColor Yellow
        throw
    }

    $url = "url"
    Enter-SeUrl $url -Driver $Driver

    $video = Find-SeElement -Driver $driver -TagName video

    $video

} catch {
    $_.Exception.Message
} finally {
    if($Null -ne $driver) {
        Stop-SeDriver -Driver $driver
        $driver = $Null
    }
}
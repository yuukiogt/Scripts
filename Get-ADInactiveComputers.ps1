$targets = Search-ADAccount -AccountInactive -TimeSpan 30 -ComputersOnly |
Where-Object {$_.Enabled -eq $True}

$targets | Out-GridView
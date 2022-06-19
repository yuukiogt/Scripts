net stop wuauserv
net stop cryptsvc
net stop bits
net stop msiserver

Move-Item -Path "C:\Windows\SoftwareDistribution" -Destination "C:\Windows\SoftwareDistribution.old" -Force
Move-Item -Path "C:\Windows\System32\catroot2" -Destination "C:\Windows\System32\catroot2.old" -Force

net start msiserver
net start bits
net start cryptsvc
net start wuauserv

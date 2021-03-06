#Get-WmiObject -Class "Win32_UserAccount" -Namespace "root\CIMV2" -Filter "LocalAccount = True" | Select-Object Name, SID | Where-Object { $_.SID -match "S-1-5-21.*?-500" }

#Get-WmiObject -Class "Win32_UserAccount" -Namespace "root\CIMV2" -Filter "LocalAccount = True" | Where-Object { $_.SID -match "S-1-5-21.*?-500" } | Select-Object Name

(Get-WmiObject -Class "Win32_UserAccount" -Namespace "root\CIMV2" -Filter "LocalAccount = True" | Where-Object { $_.SID -match "S-1-5-21.*?-500" }).Name

(Get-WmiObject -Query 'SELECT Name,SID FROM Win32_UserAccount WHERE LocalAccount = True' | Where-Object { $_.SID -match "S-1-5-21.*?-500" }).Name

(Get-WmiObject -Query 'SELECT Name,SID FROM Win32_UserAccount WHERE LocalAccount = True' | Where-Object { $_.SID -match 'S-1-5-21.*?-500' }).Name

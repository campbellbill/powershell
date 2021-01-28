<#
-c "pushd D:\Scripts; D:\Scripts\Get-VMwareSnaphots.ps1 -SendMail -FromAddr "VMwareSnaphotsReport@squaretwofinancial.com" -ToAddr "server-alerts@squaretwofinancial.com" -SmtpSvr "mail.squaretwofinancial.com""
-SendMail:$true -FromAddr "server-alerts@squaretwofinancial.com" -ToAddr bcampbell@squaretwofinancial.com,itopscorner@squaretwofinancial.com -SmtpSvr "mail.squaretwofinancial.com"
#>

powershell -c "pushd D:\Scripts; D:\Scripts\Get-VMwareSnaphots.ps1 -SendMail -FromAddr VMwareSnaphotsReport@squaretwofinancial.com -ToAddr svralts@squaretwofinancial.com -SmtpSvr mail.squaretwofinancial.com"

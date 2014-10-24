C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -command ". 'D:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell ; C:\Scripts\Get-ExchangeEnvironmentReport.ps1 -htmlreport C:\scripts\ExchangeReport.htm -SendMail:$true -MailFrom:DeanSesko@planetsesko.com -MailTo:DeanSesko@planetSesko.com -MailServer:192.168.1.216 "

exit

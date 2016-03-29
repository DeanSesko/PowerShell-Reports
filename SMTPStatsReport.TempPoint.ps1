$Localhost = $env:COMPUTERNAME
#Load Exchange PS Snapin

Write-Host "running..."

$hubs = Get-TransportServer

# Get the start date for the tracking log search
$End = Get-Date
$Start = $End.AddDays(-5)

# Get the end date for the tracking log search


$Datum = $Start.ToShortDateString()

$receive = $hubs | get-messagetrackinglog -Start $Start -End $End -EventID "RECEIVE" -ResultSize Unlimited | select Sender, RecipientCount, TotalBytes, Recipients
$send = $hubs | get-messagetrackinglog -Start $Start -End $End -EventID "SEND" -ResultSize Unlimited | select Sender, RecipientCount, TotalBytes
$mreceive = $receive | Measure-Object TotalBytes -maximum -minimum -average -sum
$msend = $send | Measure-Object TotalBytes -maximum -minimum -average -sum

$anzahl = $mreceive.count + $msend.count
$volumen = ($mreceive.sum + $msend.sum) / (1024 * 1024)

$volumen = "{0:N2}" -f $volumen + " MB"

$msendmb = $msend.sum / (1024 * 1024)
$vsend = "{0:N2}" -f $msendmb + " MB"
$bigsend = $msend.maximum / (1024 * 1024)
$avsend = $msend.average / 1024

$bigsendmb = "{0:N2}" -f $bigsend + " MB"
$avsendkb = "{0:N2}" -f $avsend + " KB"

$mreceivemb = $mreceive.sum / (1024 * 1024)
$vreceive = "{0:N2}" -f $mreceivemb + " MB"
$bigreceive = $mreceive.maximum / (1024 * 1024)
$avreceive = $mreceive.average / 1024

$bigreceivemb = "{0:N2}" -f $bigreceive + " MB"
$avreceivekb = "{0:N2}" -f $avreceive + " KB"

$computer = gc env:computername
$obj = new-object psObject

$obj | Add-Member -MemberType noteproperty -Name "Generated on server:" -Value $Computer
$obj | Add-Member -MemberType noteproperty -Name "Date :" -Value $Datum
$obj | Add-Member -MemberType noteproperty -Name "Sent mails :" -Value $msend.Count
$obj | Add-Member -MemberType noteproperty -Name "Size of sent mails:" -Value $vsend
$obj | Add-Member -MemberType noteproperty -Name "Size of biggest mail out:" -value $bigsendmb
$obj | Add-Member -MemberType noteproperty -Name "Average size out :" -value $avsendkb
$obj | Add-Member -MemberType noteproperty -Name "Quantity incoming mails :" -Value $mreceive.Count
$obj | Add-Member -MemberType noteproperty -Name "Size of received mails :" -Value $vreceive
$obj | Add-Member -MemberType noteproperty -Name "Size of biggest mail in :" -value $bigreceivemb
$obj | Add-Member -MemberType noteproperty -Name "Average size in :" -value $avreceivekb
$obj | Add-Member -MemberType noteproperty -Name "Overall quantity :" -Value $anzahl
$obj | Add-Member -MemberType noteproperty -Name "Overall size :" -Value $volumen

$out = $Datum + ";" + $msend.count + ";" + $vsend + ";" + $mreceive.count + ";" + $vreceive + ";" + $anzahl + ";" + $volumen
$out | out-file c:\scripts\daily.csv -append -encoding default

$obj = $obj -replace ("@{", "")
$obj = $obj -replace ("=", ":`t")
$obj = $obj -replace ("; ", "`n")
$obj = $obj -replace ("}", "`n")
$obj | out-file C:\Scripts\SMTPStats.txt
$MYDatabases = Get-MailBoxDatabase
$report = @()
foreach ($DB in
$MYDatabases)
{
[string]$DBPathString = $db.EDBFILEpath
# Build Database File Path
$MyDriveString =$DBPathString.substring(0,1);
$MySubString=$DBPathString.substring(0,2);
$UNCPath = “\\$($db.Server)\”
+ $MyDriveString + “$” ;
$Shell = new-object -com Shell.Application;
$UNCDBPath =  $DBPathString.Replace($MySubString,$UNCPath);
$drives = Get-WmiObject -ComputerName $DB.Server Win32_LogicalDisk |
Where-Object {$_.DeviceID -eq $MyDriveString +”:”}
foreach($drive in
$drives)
{
$size1 = $drive.size / 1GB
$size = “{0:N2}” -f
$size1
$size = $size + ” GB”
$free1 = $drive.freespace / 1GB
$free = “{0:N2}” -f $free1
$free = $free + ” GB”
}
#Get File Info
$DBFiles = get-childitem -path $UNCDBPath;
[decimal]$GBSize = ($DBFiles | Measure-Object -Property Length -Sum).Sum;
[decimal]$GBSize = $GBSize/1073741824;
$DBFILESIZE = “{0:N2}”
-f $GBSize;
$DBFILESIZe = $DBFILESIZE + “GB”;
# Output info
$MyData = New-Object Object;
$MyData | Add-Member
NoteProperty Name $DB.Name;
$MyData | Add-Member NoteProperty
Server $DB.Server;
$MyData | Add-Member NoteProperty Size $DBFILESIZE;
$MyData | Add-Member NoteProperty “Lun Size” $size;
$MyData |
Add-Member NoteProperty “Free Space” $free;
$report += $MyData
}
$report |FT -autosize
$MYDatabases = Get-Storagegroup
$report = @()
foreach ($DB in
$MYDatabases)
{
[string]$logPathString = $db.LogFolderPath
$MyDriveString=$logPathString.substring(0,1);
$MySubString=$logPathString.substring(0,2);
$UNCPath = “\\$($db.Server)\ + $MyDriveString + $ ;
$Logfoldername =$logPathString.Replace($MySubString,$UNCPath);
$Shell = new-object -com
Shell.Application
$Logfolder = $shell.namespace($Logfoldername)
$logfolderitems = $logfolder.items()
$logcount = $logfolderitems.count
$drives = Get-WmiObject -ComputerName $DB.Server Win32_LogicalDisk |
Where-Object {$_.DeviceID -eq $MyDriveString +:}
foreach($drive in
$drives)
{
$size1 = $drive.size / 1GB
$size = “{0:N2} -f $size1
$size = $size +  GB
$free1 = $drive.freespace / 1GB
$free =
“{0:N2} -f $free1
$NewLogs1 = $free / .001
$free = $free +  GB
}
$MyData = New-Object Object;
$MyData | Add-Member NoteProperty DB -value $DB.Name;
$MyData | Add-Member NoteProperty Server $DB.Server;
$MyData | Add-Member NoteProperty Logs $logcount;
$MyData | Add-Member NoteProperty Lun Size $size;
$MyData | Add-Member NoteProperty Free Space $free;
$MyData | Add-Member NoteProperty Available Logs $NewLogs1;
$report += $MyData
}
$report |FT -autosize
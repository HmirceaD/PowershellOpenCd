$disk = New-Object -ComObject IMAPI2.MsftDiscMaster2 
$disk_mover = New-Object -ComObject IMAPI2.MsftDiscRecorder2 
$disk_mover.InitializeDiscRecorder($disk) 

$disk_mover.EjectMedia() 


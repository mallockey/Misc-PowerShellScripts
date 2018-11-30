$psscriptroot = get-location | select -expandproperty path
$servers = get-content servers.txt
$driveLetters = get-content driveLetters.txt
$todaysDate = get-date -format MM-dd-yy

write-host "Getting Space info..."
function getFreeSpace{

	Param(
	$server, 
	$driverLetter
	)
	try 
	{	
	$freeSpace = Get-wmiobject -class win32_logicaldisk -computername $server  | Where-Object { $_.DeviceID -eq "$driveLetter"} | select-object -expandproperty freespace -erroraction stop
	$freeSpace = $freeSpace / 1gb
	
	}
	catch{
	$freeSpace = 0
	}

return $freeSpace
}
function getTotalSpace{

	Param(
	$server, 
	$driverLetter
	)

	try {	
	     $totalSpace = Get-wmiobject -class win32_logicaldisk -computername $server  | Where-Object { $_.DeviceID -eq "$driveLetter"} | select-object -expandproperty size -erroraction stop
	     $totalSpace = $totalSpace / 1gb
	    }
		
	catch{
	     $totalSpace = 0
	     }
	return $totalSpace
}
foreach($server in $servers){

	foreach($driveLetter in $driveLetters)
	{	
		[int]$freeSpace = getFreeSpace -driveLetter $driveLetter -server $server	
		[int]$totalSpace = getTotalSpace -driveLetter $driveLetter -server $server
		  if($freeSpace -eq 0 -or $totalSpace -eq 0)
		  {
	           write-output "$server does not have an $driveLetter drive" | out-file $psscriptroot\$todaysDate.txt -append
		  }
		  else
		  {
		  write-output "$server $driveLetter $freeSpace GBs free of $totalSpace GBs" | out-file $psscriptroot\$todaysDate.txt -append
		  }

	}
		write-output "--------------------------------------------------------" | out-file $psscriptroot\$todaysDate.txt -append

}


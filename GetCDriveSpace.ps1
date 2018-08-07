#Author: Josh Melo
#Description: Gets C drive space on each server listed in servers.txt file.
#Do not include spaces after each computer name in the text file, one PC name per line 

try{
$serversArray = get-content servers.txt -ErrorAction Stop
}
Catch{
write-host -ForeGroundColor red "Servers.txt cannot be found in directory of script."
read-host "Press enter to exit"
exit
}
write-host -foreGroundColor green "This script calculates disk space for computers."
write-host -foreGroundColor yellow "Criteria for low disk space is less than 25%"
write-host -foreGroundColor red "Criteria for very low disk space is less than 10%"

write-host "========================================================================="

[int]$lowCounter = 0
[int]$veryLowCounter = 0

start-transcript -path  "$psscriptroot\output.txt" -append
write-host "========================================================================="

	foreach ($servers in $serversArray){
		if(test-connection -ComputerName $servers -Count 1 -quiet)
		{
		try{
		#Get Total Capacity for C Drive
		$totalCapacity = get-wmiobject win32_logicaldisk -computername $servers -filter "deviceID='C:'" | foreach-object {$_.Size}  -ErrorAction stop
		$totalCapacity = $totalCapacity /1gb 
		$totalCapacity = [int]$totalCapacity 

		#Get Free Space for C drive
		$freeSpace = get-wmiobject win32_logicaldisk -computername $servers -filter "deviceID='C:'" | foreach-object {$_.FreeSpace} 
		$freeSpace = $freeSpace /1gb
		$freeSpace = [int]$freeSpace
		[int]$isGood = ($freeSpace / $totalCapacity) * 100
		$userName = get-wmiobject -computername $servers -class Win32_computersystem | select -expandproperty username

		if($isGood -le 25 -and $isGood -gt 10)
		{
		write-host -foreGroundColor yellow "Server Name:"$servers
		write-host -foreGroundColor yellow "Free Space:"$freeSpace"GBs free out of "$totalCapacity"GBs, "$isGood"% free"
		write-host -foreGroundColor yellow "User:"$userName
		write-host -foreGroundColor yellow "Status: Disk space low "
		$lowCounter++
		}
		elseif($isGood -le 10)
		{
		write-host -foreGroundColor red "Server Name:"$servers
		write-host -foreGroundColor red "Free Space:"$freeSpace"GBs free out of "$totalCapacity"GBs, "$isGood"% free"
		write-host -foreGroundColor red "User:"$userName
		write-host -foreGroundColor red "Status: Disk space very low "
		$veryLowCounter++
		}
		else
		{
		write-host -foreGroundColor green "Server Name:"$servers
		write-host -foreGroundColor green "Free Space:"$freeSpace"GBs free out of "$totalCapacity"GBs, "$isGood"% free"
		write-host -foreGroundColor green "User:"$userName
		write-host -foreGroundColor green "Status: OK"
		}
		write-host "=========================================================================" 
		}
		catch{
		if ($_.Exception.Message -like "The RPC server is unavailable.*") {$runad = 'rpc unavailable'}
		}
		}
		else{
		write-host -ForeGroundColor yellow $servers "is not responding to pings."
		}
		}
write-host -foreGroundColor green "================================Results================================="
if($lowCounter -gt 0){
write-host -foreGroundColor yellow "There are "$lowCounter "computers with low disk space."
}
else{
write-host -foreGroundColor green "There are 0 computers with low disk space."
}
if($veryLowCounter -gt 0){
write-host -foreGroundColor red "There are "$veryLowCounter "computers with low disk space."
}
else{
write-host -foreGroundColor green "There are 0 computers with very low disk space."
}
write-host "========================================================================="
stop-transcript 
Read-Host "Press enter to exit"

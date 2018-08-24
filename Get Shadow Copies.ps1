$scriptLoc = (Get-Location)
$computers = get-content -path $scriptloc\servs.txt
[int]$totalShadowCopies = 0
write-host -foreGroundColor yellow "=====================Get Shadow Copies============================"
write-host -foreGroundColor yellow "Gets Shadow Copies for each server in servers.txt"

write-host -foreGroundColor yellow "======================================================"

	foreach ($computer in $computers){
	write-host -foreGroundColor yellow "Getting Shadow Copies from:"$computer
	write-host "========================================================"	

	$dates = get-wmiobject -class win32_shadowcopy -computerName $computer | select-object -expandProperty installdate
	$dates = $dates | sort-object 

		foreach ($date in $dates){
		$totalShadowCopies++
		$date = "$date".Insert(4,"-")
		$date = "$date".Insert(7,"-")
		$date = "$date".Insert(10,"-")
		$date = $date.Substring(0,$date.length - 15)
		[int]$hours = $date.substring($date.length - 2)
		$date = $date.Substring(0,$date.length - 3)
		$date = get-date -format 'MM/dd/yyyy' $date
		
			if($hours -gt 12){
				$hours = $hours - 12
				$hours = $hours
				$time = "PM"
			}
			elseif($hours -eq 0){
				
				$hours = 12
				$time = "PM"
			}
			else{
				$hours = $hours
				$time = "AM"
			}
		
		write-host -foreGroundColor green $date $hours$time

		}
	write-host "========================================================"	
	write-host -foreGroundColor yellow "Total Shadow copies for $computer"$totalShadowCopies
	$totalShadowCopies = 0
	write-host "========================================================"	
}
read-host "Press Enter to exit"

$scriptLoc = (Get-Location)
$computers = get-content -path $scriptloc\servers.txt
[int]$totalShadowCopies = 0
write-host -foreGroundColor yellow "=====================Get Shadow Copies============================"
write-host -foreGroundColor yellow "Gets Shadow Copies for each server in servers.txt"
start-sleep -Milliseconds 1000
foreach ($computer in $computers){
    if(test-connection $computer -count 1 -quiet){
    $dates = get-wmiobject -class win32_shadowcopy -computerName $computer | select-object -expandProperty installdate
	if($dates -eq $null){
	write-host -foreGroundColor red $computer "does not have any shadow copies."
	write-host -foreGroundcolor yellow  "========================================================"
	continue
	}
	$dates = $dates | sort-object 
	write-host -foreGroundcolor yellow  "========================================================"	
	write-host -foregroundColor yellow Oldest - $computer
	write-host -foreGroundcolor yellow  "========================================================"	
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
		write-host -foreGroundcolor yellow  "========================================================"	
		write-host -foregroundColor yellow Newest - End of $computer
		write-host -foreGroundcolor yellow  "========================================================"	
		write-host -foreGroundColor yellow "Total Shadow copies for $computer"$totalShadowCopies
		$totalShadowCopies = 0
		write-host -foreGroundColor yellow  "========================================================"	
	    }
    else{
	write-host -foregroundcolor red $computer "is not online."
	write-host -foreGroundcolor yellow  "========================================================"
    }
	
}
write-host -foreGroundColor green "Completed. Press Enter to exit"

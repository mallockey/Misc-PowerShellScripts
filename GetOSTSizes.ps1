$computersArray = get-content computers.txt
    foreach($computer in $computersArray){
    $ostArray = Get-ChildItem -Path "\\$computer\c$\users\*\appdata\local\microsoft\outlook\*.ost"
    write-host $computer
    $totalOSTsOnPC = 0
        foreach($ost in $ostArray){	
	    $ostLength = $ost.length
	    $ostLengthInGB = $ostLength / 1gb
	    $ostLengthInGb = [int]$ostLengthInGb
	    $ostTotalSum += $ost.length
	    $ostTotalSumInGb = $ostTotalSum / 1gb
	    $ostTotalSumInGb = [int]$ostTotalSumInGb
	    $totalOSTsOnPC++
	    write-host $ost.name "|" $ostLengthInGB"GBs"
	}
	    $ostTotalSum = $ostTotalSum / 1gb
	    $ostTotalSum = [int]$ostTotalSum
	    
	    if($ostTotalSum -gt 30){
	        write-host -foregroundColor yellow "Total Size all OSTS:$ostTotalSum GBs "Warning, a lot of space is being used""
	    }
	    else{
	        write-host Total Size all OSTS: $ostTotalSum "GBs"
	    }
    write-host "Total OSTs:"$totalOSTsOnPC			
    write-host "============================================================"			
    }

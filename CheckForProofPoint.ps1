$scriptLoc = (Get-Location)
$computersArray = get-content -path $scriptLoc\computers.txt

	foreach($computer in $computersArray){
	$PP32 = "\\$computer\c$\Program Files (x86)\Proofpoint"
	$PP64 = "\\$computer\c$\Program Files\Proofpoint"
	
	if(test-connection $computer -count 1 -quiet){
	
			if(test-path -path $PP32){
			write-host -foreGroundColor green $computer "has Proofpoint 32 installed."
			}
			elseif(test-path -path $PP64){
			write-host -foreGroundColor green $computer "has Proofpoint 64 installed."
			}
			else{
			write-host "$computer - not installed"
			}
	}
	else{
	write-host $computer "is offline"
	}
			
}
write-host -foregroundcolor green "Finished"
read-host

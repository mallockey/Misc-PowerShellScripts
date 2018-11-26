<#
Author: Josh Melo
Last Updated: 11/25/18
This sript will check that Proofpoint outlook plugin is installed on each PC in a given workstation OU.
It will only check that Proofpoint folder exists on C:\Program Files and C:\Program Files (x86)

#>
try{
import-module activedirectory -erroraction stop
}
catch{
write-host -foreGroundColor red "Run this from a DC"
read-host
exit
}

$ou = read-host "Enter the OU where the PCs are"

try
{
    $computersArray = get-adcomputer -filter * -searchbase $ou| select -expandproperty name
}
catch
{
    write-host -foreGroundColor red "OU not correct please verify OU and rerun."
    read-host 
    exit
}

	foreach($computer in $computersArray){
	$PP32 = "\\$computer\c$\Program Files (x86)\Proofpoint"
	$PP64 = "\\$computer\c$\Program Files\Proofpoint"
	
	if(test-connection $computer -count 1 -quiet){
	
			if(test-path -path $PP32){
			write-host -foreGroundColor green $computer "has Proofpoint 32 bit installed."
			}
			elseif(test-path -path $PP64){
			write-host -foreGroundColor green $computer "has Proofpoint 64 bit installed."
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

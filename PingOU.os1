try{
import-module activedirectory -erroraction stop
}
catch{
write-host -foreGroundColor red "Run this from a DC"
read-host
exit
}
$currentDir = "$psscriptroot"
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
$offlinePCs = 0
foreach($computer in $computersArray){
	
	if(test-connection $computer -count 1 -quiet)
	{
	write-host -foregroundcolor green $computer "is online"
	}
	else{
	write-host -foregroundcolor yellow $computer "is not online"
	$offlinePCs++
	$offline = "Yes"
	$wrapper = New-Object PSObject -Property @{ ComputerName = $computer;}
	$wrapper | add-member NoteProperty Offline? $offline 

	
	Export-Csv -InputObject $wrapper -Path $currentDir"\OfflinePCs.csv" -NoTypeInformation -Append
	}

}
write-host "----------------------------------------------"
write-host -foregroundColor yellow PCs Offline: $offlinePCs
write-host -foregroundColor green "Offline PCs exported to CSV at : $psscriptroot"
write-host -foregroundcolor green "Finished"
read-host

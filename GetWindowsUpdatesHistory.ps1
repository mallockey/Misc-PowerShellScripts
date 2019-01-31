$scriptLoc = (Get-Location)
write-host "This script gets Windows Updates from workstations"

$computersArray = get-content -path  $scriptLoc\Computers.txt
write-host "Working, please wait..."
    foreach ($computers in $computersArray){
	if(test-connection -computerName $computers -Quiet -count 1)
	{
	write-host $computers "is online. writing to WindowsUpdateReview.txt..."
	get-hotfix -computerName $computers | select  pscomputername, hotfixid, installedon | 
	sort installedon  | export-csv $scriptLoc\WindowsUpdates.csv -append -notypeinformation
	}
	else{
	write-host $computers "is not online"
	}
    }

read-host "Results located at $scriptLoc"

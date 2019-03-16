import-module activedirectory
$ou = ""
$workstations = get-adcomputer -filter * -searchbase $ou | select -expandproperty name
foreach($workstation in $workstations){
    if(test-connection -count 1 $workstation -quiet){
        try{
	    $currentLoggedOnUser = get-wmiobject -class win32_computersystem -computerName $workstation | select -expandproperty username
            $tempIndex = $currentLoggedOnUser.IndexOf("\") + 1
	    $currentLoggedOnUser = $currentLoggedOnUser.SubString($tempIndex)
	    $fullUserName = get-aduser -identity $currentLoggedOnUser | select -expandproperty name
	    write-host $workstation $fullUserName
	 }
	 catch{
	     $currentLoggedOnUser = "No User Logged On"
	     write-host $currentLoggedOnUser
	     continue
	 }
	Set-ADComputer -Identity $workstation -Description $fullUserName
    }
}


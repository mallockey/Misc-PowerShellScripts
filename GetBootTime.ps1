import-module ActiveDirectory
$ou = read-host "Enter the OU where the PCs are"
try{
$computersArray = get-adcomputer -filter * -searchbase $ou| select -expandproperty name
}
catch{
write-host -foreGroundColor red "OU not correct please verify OU and rerun."
read-host 
exit
}

$scriptLoc = (Get-Location)
foreach($computer in $computersArray){
    if(test-connection $computer -quiet -count 1){
    try{
    $userName = get-wmiobject -computername $computer -class Win32_computersystem | select -expandproperty username
    if($userName -eq $null){
    $userName = "No User logged on"
    }
    }
    catch{
    write-host "User not logged on"
    }
    $today = get-date
    $booted = get-wmiobject -class win32_operatingsystem -computerName $computer
    $lastBoot = $booted.converttodatetime($booted.lastbootuptime)
    [int]$days = new-timespan -start $lastBoot -end $today | select -expandproperty days
    $hours = new-timespan -start $lastBoot -end $today | select -expandproperty hours
    $minutes = new-timespan -start $lastBoot -end $today | select -expandproperty minutes
    write-host "Computer Name:"$computer 
    write-host "User Name:"$userName 
        if($days -gt 30){
	write-host -foregroundColor red "Uptime:"$days "days,"$hours "hours and"$minutes" minutes. Probably need a reboot"
	}
	else{
        write-host -foregroundColor green "Uptime:"$days "days,"$hours "hours and"$minutes" minutes"
	}
    write-host "========================================================================="
    }
    else{
    write-host -foreGroundColor red $computer "is not online."
    write-host "========================================================================="
    }
}
read-host "Press Enter to exit"

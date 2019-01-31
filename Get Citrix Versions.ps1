$scriptLoc = (Get-Location)
$computersArray = get-content -path $scriptLoc\computers.txt
   foreach($computer in $computersArray){
   $cdviewer = "\\$computer\c$\program files\citrix\ica client\cdviewer.exe"
   $cdviewer64 = "\\$computer\c$\program files (x86)\citrix\ica client\cdviewer.exe"
	if(test-path $cdviewer){
	$citrixver = (get-command $cdviewer).FileVersionInfo.Fileversion
	write-host $computer "-" $citrixver
	}
	elseif(test-path $cdviewer64){
	$citrixver = (get-command $cdviewer64).FileVersionInfo.Fileversion
	write-host $computer "-" $citrixver
	}
	else{
	write-host "$computer - not installed"
	}
    }

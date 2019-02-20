Param(
[Parameter(Mandatory=$true)]
$PCConvention,
$findUser
)
import-module activedirectory
$computers = Get-ADComputer -Filter * | Where-Object {$_.Name -like "*$PCConvention*"} | select -ExpandProperty name
$PCsLoggedInto =""
foreach($computer in $computers){

	if(test-connection -count 1 -quiet $computer){
		try{
		$currentUser = get-wmiobject -class win32_computersystem -computername $computer | select -expandproperty username
		}
		catch{
			write-progress -Activity "Checking PCs" -currentOperation "Failed to get data from $computer"
		}
		Write-Progress -Activity "Checking PCs" -CurrentOperation "Current PC: $computer"
		if($currentUser -like "*$findUser*"){
			$PCsLoggedInto += "$computer "
		}
		else{
		continue
		}
	}
	else{
	Write-Progress -Activity "Checking PCs" -CurrentOperation "Current PC: $computer is offline"
	}
}
write-host "$findUser was logged into the below PCs:"
write-host $PCsLoggedInto

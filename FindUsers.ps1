Param(
[Parameter(Mandatory=$true)]
$PCConvention
)
import-module activedirectory
$computers = Get-ADComputer -Filter * | Where-Object {$_.Name -like "*$PCConvention*"} | select -ExpandProperty name

foreach($computer in $computers){

	if(test-connection -count 1 -quiet $computer){
		try{
			$currentUser = get-wmiobject -class win32_computersystem -computername $computer -erroraction stop | select -expandproperty username
				if($currentUser -eq $null){
					$currentUser = "No User Logged on"
				}
			write-host "$computer - $currentUser"
		}
		catch{
			write-progress -Activity "Checking PCs" -currentOperation "Failed to get data from $computer"
		}
	}
	else{
	Write-Progress -Activity "Checking PCs" -CurrentOperation "Current PC: $computer is offline"
	}
}

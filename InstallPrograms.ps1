write-host -foreGroundColor yellow("This script will attempt to install all setup files in a given directory")
write-host ----------------------------------------------------------
function getArgument($program){

	$program = (get-command $program).FileVersionInfo.Filedescription
		if($program -match 'Microsoft Setup Bootstrapper'){
		$program = "/config config.xml"
		return $program
	}
	if($program -match '.msi'){
	$program = "/qn"
	return $program
	}
	elseif($program -match '.exe'){
	$program = "/silent"
	return $program
	}
	else{
	return $program
	}

}
function successText($output)
{
	$output = write-host -ForeGroundColor green "SUCCESS:"$output
	return $output
}

function failText($output)
{
	$output = write-host -ForeGroundColor red "ERROR:"$output
	return $output
}
function infoText($output)
{
	$output = write-host -foreGroundColor yellow "INFO:"$output
	return $output
}
function isEmptyString($string)
{

	if([string]::IsNullOrWhiteSpace($string))
		{
		return $true
		}
		else
		{
		return $false
		}

}
function testPath($path)
{
 $path = test-path $path
	 if($path -eq $true)
	 {
	 return $true
	 }
	 else
	 {
	 return $false
	 }

}

$programsDir = read-host "Enter the path of the install files are include \ at end (Ex. c:\kits\)"
try{
$validatePath = testPath($programsDir)
$validateNull = isEmptyString($programsDir)
}
catch{
failText("Path is either not valid or string is empty")
failText("Please rerun script")
read-host
exit
}
if($programsDir.EndsWith("\") -eq $false){
	failText("$programsDir does not end with \")
	read-host
	exit
}

$programsFoundInFolder= get-childitem $programsDir -name
$totalProgramsInDirectory = $programsFoundInFolder.count
$indexOfPrograms = 1
$invalidProgramsCounter = 0
$installFails = 0
$arrayOfValidPrograms = [System.Collections.ArrayList]@()

		for($i=0; $i -lt $totalProgramsInDirectory; $i++){
		$currentProgram = $programsFoundInFolder[$i]
		
			if($currentProgram -match '.exe' -or $currentProgram -match '.msi'){
				
				$array = $arrayOfValidPrograms.add("$programsDir$currentProgram")
				write-host $indexOfPrograms")" $programsFoundInFolder[$i]
				
				$indexOfPrograms++
			}
			else{
				$invalidProgramsCounter++
			}
			
		}
write-host ----------------------------------------------------------
infoText("$invalidProgramsCounter files have been omitted because they are not valid install files")
infoText("This will install the above programs on your PC")
write-host "Press Enter to continue:"
read-host

for($i=0; $i -lt $arrayOfValidPrograms.count; $i++){
	$currentArgument = getArgument($arrayOfValidPrograms[$i])
	write-host $arrayOfValidPrograms[$i] $currentArgument
	write-host -foreGroundColor green "Installing "$arrayOfValidPrograms[$i]"...Please wait"
	$installer = start-process $arrayOfValidPrograms[$i] -argumentlist $currentArgument -wait -passthru
	write-host -------------------------------------------
		if($installer.ExitCode -eq 0)
				{
				successText($arrayOfValidPrograms[$i] + " installed successfully!")
				}
			else{
				failText ($arrayOfValidPrograms[$i] +  " did not install successfully, Error Code is: $($installer.ExitCode)")
				$installFails++
				}
				
}

write-host -----------------------------------------------------------
write-host -foreGroundColor green "Finished."
infoText("$installFails programs failed to install.")
read-host

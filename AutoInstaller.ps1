<#
Author: Josh Melo
Last Updated: 11/25/18
This script wil attempt to install all the installer files in a given a directory. It is hardcoded to look for a config.xml file
for any Volume Licensed office install. It also has the correct argument for Google Chrome. Most of these checks are based on the file name
matching the product so please don't rename the setup files names.
#>
write-host -foreGroundColor cyan("============================Auto Installer============================")
write-host -foreGroundColor cyan("This script will attempt to install all setup files in a given directory")
write-host -foreGroundColor cyan ----------------------------------------------------------
function getArgument(){
Param(
$program
)
    if((get-command $program).FileVersionInfo.Filedescription -match 'Microsoft Setup Bootstrapper'){
    $program = "/config config.xml"
    return $program
    }
    elseif($program -match 'ninite'){
    $program = $null
    return $program
    }
    elseif((get-command $program).FileVersionInfo.Filedescription -match 'Google'){
    $program = "/silent /install"
    return $program
    }
    elseif((get-command $program).FileVersionInfo.Filedescription -match 'Microsoft Office'){
    $program = $null
    return $program
    }
    elseif($program -match '.msi'){
    $program = "/qn"
    return $program
    }
    elseif($program -match '.exe'){
    $program = "/silent"
    return $program
    }
    else{
    $program = $null
    return $program
    }
}
function successText($output){
$output = write-host -ForeGroundColor green "SUCCESS:"$output
return $output
}
function failText($output){
$output = write-host -ForeGroundColor red "ERROR:"$output
return $output
}
function infoText($output){
$output = write-host -foreGroundColor yellow "INFO:"$output
return $output
}
function isEmptyString($string){
    if([string]::IsNullOrWhiteSpace($string)){
    return $true
    }
    else{
    return $false
    }
}
function testPath($path){
$path = test-path $path
    if($path -eq $true){
    return $true
    }
    else{
    return $false
    }
}
$currentDir = "$psscriptroot"
$programsDir = read-host "Enter the path of the install files are(Ex. c:\kits\)"
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
$currentDir = "$psscriptroot"
copy-item "$psscriptroot\config.xml" -destination $programsDir
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
write-host -foreGroundColor green "Installing:"$arrayOfValidPrograms[$i]
    if($currentArgument -eq $null){
    $installer = start-process $arrayOfValidPrograms[$i] -wait -passthru		
        if($installer.ExitCode -eq 0){
	successText($arrayOfValidPrograms[$i] + " installed successfully!")
	}
	else{
	failText ($arrayOfValidPrograms[$i] +  " was unsuccessful.")
	write-host -foreGroundColor red "ERROR:"$installer.ExitCode
	$installFails++
	}
    }
    else{
    $installer = start-process $arrayOfValidPrograms[$i] -argumentlist $currentArgument -wait -passthru
        if($installer.ExitCode -eq 0){
	successText($arrayOfValidPrograms[$i] + " installed successfully!")
	}
	else{
	failText ($arrayOfValidPrograms[$i] +  " was unsuccessful.")
	write-host -foreGroundColor red "ERROR:"$installer.ExitCode
	$installFails++
	}
    }
    write-host -------------------------------------------
}
infoText("$installFails programs failed to install.")
read-host

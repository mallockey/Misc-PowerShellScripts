function getArgument($program){

if($program -match '.msi'){
$program = "/qn"
return $program
}
elseif($program -match '.exe'){
$program = "/s"
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

$programsDir = "C:\kits\"
$programsFoundInFolder= get-childitem "c:\kits\" -name
$totalProgramsInDirectory = $programsFoundInFolder.count
$indexOfPrograms = 1
$programsArrayOfPaths = New-Object -TypeName 'object[]' -ArgumentList $totalProgramsInDirectory

for($i=0; $i -lt $totalProgramsInDirectory; $i++){
$currentProgram = $programsFoundInFolder[$i]
if($currentProgram -match '.exe' -or $currentProgram -match '.msi'){
$programsArrayOfPaths[$i] = "$programsDir$currentProgram"
write-host $indexOfPrograms")" $programsFoundInFolder[$i]
$indexOfPrograms++
}
else{
}
}
write-host ----------------------------------------------------------

infoText("This will install the  above programs on your PC")

read-host

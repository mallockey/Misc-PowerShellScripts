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
$input = ""
while($input -ne "stop"){
$input = read-host "Please enter input"
if(isEmptyString($input))
{
write-host "String is empty"
}
else{
	
	write-host "String is NOT Empty"
	if(testPath($input)){
	write-host "Path is valid"
	}
	else{
	write-host "Path is NOT valid"
	}
	
}

}

function getArgument($argument){

	switch($argument){
	.msi 
	{
		return $argument = "/quiet"
	}
	.exe
	{
		return $argument = "/s"
	}
	
	}

}

$msi = ".exe"
getArgument($msi)

read-host "Done, press enter"

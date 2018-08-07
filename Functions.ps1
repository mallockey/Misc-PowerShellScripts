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
$input = "C:\users"

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


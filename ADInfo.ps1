import-module ActiveDirectory
$currentDir = "$psscriptroot"
$computersArray = get-content computers.txt
$arrayOfInfo = [System.Collections.ArrayList]@()
foreach($computer in $computersArray){
  try{
  $ADInfo = get-adcomputer -identity $computer -properties *
  }
  catch{

    $description = "NOT IN AD"
    $lastlogondate = "N/A"
    $isEnabled = "N/A"

    $tempObj = New-Object -TypeName PSObject  
    $tempObj | Add-Member -MemberType NoteProperty -Name "PCName" -Value $computer
    $tempObj | Add-Member -MemberType NoteProperty -Name "Description" -Value $description
    $tempObj | Add-Member -MemberType NoteProperty -Name "LastLogon" -Value $lastlogondate
    $tempObj | Add-Member -MemberType NoteProperty -Name "Enabled" -Value $isEnabled
    $arrayOfInfo.Add($tempObj) | out-null 
    continue
  }

	$description = $ADInfo.Description
	$lastlogondate = $ADInfo.LastLogonDate
	$isEnabled = $ADInfo.enabled 

	$tempObj = New-Object -TypeName PSObject  
	$tempObj | Add-Member -MemberType NoteProperty -Name "PCName" -Value $computer
	$tempObj | Add-Member -MemberType NoteProperty -Name "Description" -Value $description
	$tempObj | Add-Member -MemberType NoteProperty -Name "LastLogon" -Value $lastlogondate
	$tempObj | Add-Member -MemberType NoteProperty -Name "Enabled" -Value $isEnabled
	$arrayOfInfo.Add($tempObj) | out-null 
	write-host $computers $description "Last Logon Date:"$lastLogonDate "IsEnabled" $isEnabled
	
}

$arrayOfInfo | format-table
$arrayOfInfo | export-csv $psscriptroot"\Results.csv" -noTypeInformation

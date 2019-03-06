import-module ActiveDirectory
$currentDir = "$psscriptroot"
$computersArray = get-adcomputer -filter * | select -expandproperty name
$arrayOfInfo = [System.Collections.ArrayList]@()
foreach($computer in $computersArray){
 
    $ADInfo = get-adcomputer -identity $computer -properties *
    write-progress -Activity "Collecting Data" -Status "Current PC: $computer"	

    $description = $ADInfo.Description
    $lastlogondate = $ADInfo.LastLogonDate
    $isEnabled = $ADInfo.enabled 
    
    $tempObj = New-Object -TypeName PSObject  
    $tempObj | Add-Member -MemberType NoteProperty -Name "PCName" -Value $computer
    $tempObj | Add-Member -MemberType NoteProperty -Name "Description" -Value $description
    $tempObj | Add-Member -MemberType NoteProperty -Name "LastLogon" -Value $lastlogondate
    $tempObj | Add-Member -MemberType NoteProperty -Name "Enabled" -Value $isEnabled
    $arrayOfInfo.Add($tempObj) | out-null 
    
}

$arrayOfInfo | format-table
$arrayOfInfo | export-csv $psscriptroot"\StalePCs.csv" -noTypeInformation

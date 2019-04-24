$OU = ""
$PCS = Get-ADComputer -filter * -searchbase $OU -properties *
$arrayOfInfo = [System.Collections.ArrayList]@()

foreach($PC in $PCS){
    $tempObj = New-Object -TypeName PSObject
    $DN = $PC.DistinguishedName
    $PCName = $PC.Name
    $Description = $PC.Description
    $recoveryPassword = Get-ADObject -Filter {objectclass -eq 'msFVE-RecoveryInformation'} -SearchBase $DN  | select msfve-recoverypassword
    if($recoveryPassword -ne $null){
        $recoveryPassword = "Good"
    }
    else{
	 $recoveryPassword = "NOT BACKED UP"
    }
    $tempObj | Add-Member -MemberType NoteProperty -Name PCName -Value $PCName
    $tempObj | Add-Member -MemberType NoteProperty -Name Description -Value $Description
    $tempObj | Add-Member -MemberType NoteProperty -Name BackedUpStatus -Value $recoveryPassword
    $arrayOfInfo.Add($tempObj) | out-null 
}

$arrayOfInfo | export-csv "C:\Kits\BitlockerStatus.csv" -noTypeInformation -Append

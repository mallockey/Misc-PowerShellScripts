import-module ActiveDirectory
$ou = read-host "Enter the OU where the PCs are"
$currentDir = "$psscriptroot"
try{
$computersArray = get-adcomputer -filter * -searchbase $ou| select -expandproperty name
}
catch{
write-host -foreGroundColor red "OU not correct please verify OU and rerun."
read-host 
exit
}
foreach($computers in $computersArray){
$lastLogonDate = get-adcomputer -identity $computers -properties * | select-object -ExpandProperty lastlogondate 
write-host $computers "Last Logon Date:"$lastLogonDate
$wrapper = New-Object PSObject -Property @{ ComputerName = $computers; LastLogon = $lastLogonDate }
Export-Csv -InputObject $wrapper -Path $currentDir"\LastLogons.csv" -NoTypeInformation -Append
}
write-host -foreGroundColor green "Success! LastDates.CSV is located where the script was run from. Press Enter to exit"
read-host

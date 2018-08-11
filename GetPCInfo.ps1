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

    if(Test-Connection -computerName $computers -count 1 -quiet ){
    write-host -foreGroundColor green "Getting info from:" $computers

    $userName = get-wmiobject -computername $computers -class Win32_computersystem | select -expandproperty username
        if($userName -eq $null){
        $userName = "No User logged on"
        }
    $processor = Get-WmiObject -computerName $computers -class Win32_Processor | select -expandproperty name
    $ram = Get-WmiObject -ComputerName $computers -class Win32_computersystem | select -ExpandProperty TotalPhysicalMemory
    $ram = $ram / 1gb
    $ram = [int]$ram
    $serialNumber = Get-WmiObject -computerName $computers -Class Win32_Bios | select -expandproperty serialnumber
        if($serialNumber -eq $null){
            $serialNumber = "N/A"
        }
    $videoCard = get-wmiobject win32_videocontroller -computername $computers | select -expandProperty Description
        if($videoCard -like "*System*"){
            $videoCard = "N/A"
        }

    $os = get-wmiobject Win32_OperatingSystem -computerName $computers
    $osType = $os.caption

    $osArc = Get-WmiObject Win32_OperatingSystem -ComputerName $computers | select -ExpandProperty osarchitecture
    $wrapper = New-Object PSObject -Property @{ ComputerName = $computers;}
    $wrapper | add-member NoteProperty User $username
    $wrapper | add-member NoteProperty OS $osType
    $wrapper | add-member NoteProperty OSArch $osArc
    $wrapper | add-member NoteProperty Processor $processor
    $wrapper | add-member NoteProperty RAM $ram"GBs"
    $wrapper | add-member NoteProperty VideoCard $videoCard
    $wrapper | add-member NoteProperty SerialNumber $serialNumber
                                                
    Export-Csv -InputObject $wrapper -Path $currentDir"\PCInfo.csv" -NoTypeInformation -Append

    }
    else{
        write-host -foreGroundColor red $computers" is not online."
    }

}
write-host -foreGroundColor green "Finished. PCInfo.CSV is located where the script was run from. Press Enter to exit"
read-host

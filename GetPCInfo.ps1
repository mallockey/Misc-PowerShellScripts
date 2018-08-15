#Author: Josh Melo 8/14/18
#Run this from a domain controller. Use Attribute Editor to find Distinguished name of OU and enter into script.
#This gets the following field for each PC in the OU:
#Computer Name, Logged on User, OS, OS Architecture, Processor, Total RAM, VideoCards, and Serial Number

import-module ActiveDirectory
$ErrorActionPreference = "stop"
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

$todayDate = get-date
foreach($computers in $computersArray){

    if(Test-Connection -computerName $computers -count 1 -quiet ){

    write-host -foreGroundColor green "Getting info from:" $computers
try{
$userName = get-wmiobject -computername $computers -class Win32_computersystem -erroraction stop | select -expandproperty username
if($userName -eq $null){
$userName = "No User logged on"
}

$processor = Get-WmiObject -computerName $computers -class Win32_Processor | select -expandproperty name -erroraction stop
$ram = Get-WmiObject -ComputerName $computers -class Win32_computersystem | select -ExpandProperty TotalPhysicalMemory
$ram = $ram / 1gb
$ram = [int]$ram
$serialNumber = Get-WmiObject -computerName $computers -Class Win32_Bios | select -expandproperty serialnumber
if($serialNumber -eq $null){
$serialNumber = "N/A"
}

$manufacturer = get-wmiobject -class win32_computersystem -computerName $computers | select -expandproperty manufacturer
$model = get-wmiobject -class win32_computersystem -computerName $computers | select -expandproperty model

##############################################VideoCard############################################################################
[string]$numCards = get-wmiobject -class Win32_VideoController -computername $computers | select -expandproperty deviceid

if($numCards -like "VideoController1")
{
[int]$numCards = 1
}
elseif($numCards -like "VideoController1 VideoController2")
{
[int]$numCards = 2
}
elseif($numCards -like "VideoController1 VideoController2 VideoController3")
{
[int]$numCards = 3
}
elseif($numCards -like "VideoController1 VideoController2 VideoController3 VideoController4")
{
[int]$numCards = 4
}

for($i=1;  $i -le $numCards; $i++){
$videoCardString = "videocontroller" + "$i"
$videoCards = Get-wmiobject -class Win32_VideoController -computername $computers | where-object {$_.DeviceID -eq $videocardString}
[string]$videoCardName = get-wmiobject win32_videocontroller -computername $computers  | select -expandProperty Name
}
##############################################VideoCard############################################################################
$os = get-wmiobject Win32_OperatingSystem -computerName $computers
$osType = $os.caption
$osArc = Get-WmiObject Win32_OperatingSystem -ComputerName $computers | select -ExpandProperty osarchitecture
$wrapper = New-Object PSObject -Property @{ ComputerName = $computers;}
$wrapper | add-member NoteProperty User $username
$wrapper | add-member NoteProperty Manufacturer $manufacturer
$wrapper | add-member NoteProperty Model $model
$wrapper | add-member NoteProperty OS $osType
$wrapper | add-member NoteProperty OSArch $osArc
$wrapper | add-member NoteProperty Processor $processor
$wrapper | add-member NoteProperty RAM $ram"GBs"
$wrapper | add-member NoteProperty VideoCard $videoCardName
$wrapper | add-member NoteProperty SerialNumber $serialNumber
$wrapper | add-member NoteProperty RetrievedOn $todayDate
Export-Csv -InputObject $wrapper -Path $currentDir"\PCInfo.csv" -NoTypeInformation -Append

}
catch
{
write-host -foreGroundColor yellow $computers" is online but not responding. Likely RPC service is not running or is being blocked by Windows Firewall."
}
}

else
    {
        write-host -foreGroundColor red $computers" is not online."
    }
}

write-host -foreGroundColor green "Finished. PCInfo.CSV is located where the script was run from. Press Enter to exit"
read-host

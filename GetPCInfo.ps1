#Author: Josh Melo 8/14/18
#Last updated on 9/19/18. Changed allowing PCs to be entered into the CSV if they are not online.
import-module ActiveDirectory
write-host "======================================Get Info========================================================"
write-host -foreGroundColor yellow "This gets the following field for each PC in the OU:"
write-host -foreGroundColor green "Computer Name, Logged on User, OS, OS Architecture, Manufacturer, Model Processor"
write-host -foreGroundColor green "Total RAM, VideoCards, Serial Number and when it was retrieved on."
write-host -foreGroundColor yellow "It will then store the info in PCInventory.csv where ever the script was run from."
write-host -foreGroundColor yellow "Find the OU in AD by right clicking on the OU, go to Properties"
write-host -foregroundColor yellow "Attribute Editor -> Distinguished name"
write-host "======================================================================================================="

$ErrorActionPreference = "stop"
$ou = read-host "Enter the OU where the PCs are"
$currentDir = "$psscriptroot"

try
{
$computersArray = get-adcomputer -filter * -searchbase $ou| select -expandproperty name
}
catch
{
write-host -foreGroundColor red "OU not correct please verify OU and rerun."
read-host 
exit
}
    foreach($computers in $computersArray)
    {
    $retrievedOn = get-date
    $wrapper = New-Object PSObject -Property @{ ComputerName = $computers;}
        if(Test-Connection -computerName $computers -count 1 -quiet )
	{
	try
	{
	 $userName = get-wmiobject -computername $computers -class Win32_computersystem -erroraction stop | select -expandproperty username
             if($userName -eq $null)
	     {
	     $userName = "No User logged on"
	     }
	 }
	catch
	{
	write-host -foreGroundColor red "Error getting info from:"$computers
	continue
	}
	write-host -foreGroundColor green "Getting info from:" $computers
	$processor = Get-WmiObject -computerName $computers -class Win32_Processor | select -expandproperty name -erroraction stop
	$ram = Get-WmiObject -ComputerName $computers -class Win32_computersystem | select -ExpandProperty TotalPhysicalMemory
	$ram = $ram / 1gb
	$ram = [int]$ram
	$serialNumber = Get-WmiObject -computerName $computers -Class Win32_Bios | select -expandproperty serialnumber
	    if($serialNumber -eq $null){
	    $serialNumber = "N/A"
	    }			
	$manufacturer = get-wmiobject -class win32_computersystem -computerName $computers  | select -expandproperty manufacturer
	$model = get-wmiobject -class win32_computersystem -computerName $computers | select -expandproperty model
	$os = get-wmiobject Win32_OperatingSystem -computerName $computers
	$osType = $os.caption
	$osArc = Get-WmiObject Win32_OperatingSystem -ComputerName $computers | select -ExpandProperty osarchitecture

	##############VideoCard###################
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
	#########EndOfVideoCard##################
	$wrapper | add-member NoteProperty User $username
	$wrapper | add-member NoteProperty OS $osType
	$wrapper | add-member NoteProperty OSArch $osArc
	$wrapper | add-member NoteProperty Manufacturer $manufacturer
	$wrapper | add-member NoteProperty Model $model
	$wrapper | add-member NoteProperty Processor $processor
	$wrapper | add-member NoteProperty RAM $ram"GBs"
	$wrapper | add-member NoteProperty VideoCard $videoCardName
	$wrapper | add-member NoteProperty SerialNumber $serialNumber
	$wrapper | add-member NoteProperty RetrievedOn $retrievedOn
	Export-Csv -InputObject $wrapper -Path $currentDir"\PCInventory.csv" -NoTypeInformation -Append
		
	}
	else
	{
	write-host -foregroundcolor yellow $computers "is offline"
	}
    }
write-host -foregroundColor green "Completed. PCInventory.csv is located: $currentDir. Press enter to exit:"
read-host

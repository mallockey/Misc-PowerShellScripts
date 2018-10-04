#Author: Josh Melo 8/14/18
#Use for a single PC.
#This gets the following field for each PC in the OU:
#Computer Name, Logged on User, OS, OS Architecture, Processor, Total RAM, VideoCards, and Serial Number

$todaysDate = get-date
$hostName = hostname
    write-host -foreGroundColor green "Getting info from $hostName"
    $currentDir = "$psscriptroot"
$userName = get-wmiobject -class Win32_computersystem -erroraction stop | select -expandproperty username
    if($userName -eq $null){
    $userName = "No User logged on"
    }

$processor = Get-WmiObject -class Win32_Processor | select -expandproperty name -erroraction stop
$ram = Get-WmiObject  -class Win32_computersystem | select -ExpandProperty TotalPhysicalMemory
$ram = $ram / 1gb
$ram = [int]$ram
$serialNumber = Get-WmiObject  -Class Win32_Bios | select -expandproperty serialnumber

    if($serialNumber -eq $null){
    $serialNumber = "N/A"
    }
##############################################VideoCard############################################################################
[string]$numCards = get-wmiobject -class Win32_VideoController | select -expandproperty deviceid

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
    $videoCards = Get-wmiobject -class Win32_VideoController | where-object {$_.DeviceID -eq $videocardString}
    [string]$videoCardName = get-wmiobject win32_videocontroller  | select -expandProperty Name
    }
##############################################VideoCard############################################################################
$os = get-wmiobject Win32_OperatingSystem 
$osType = $os.caption
$osArc = Get-WmiObject Win32_OperatingSystem | select -ExpandProperty osarchitecture
$manufacturer = get-wmiobject -class win32_computersystem  | select -expandproperty manufacturer
$model = get-wmiobject -class win32_computersystem | select -expandproperty model

$wrapper = New-Object PSObject -Property @{ ComputerName = $hostName;}
$wrapper | add-member NoteProperty User $username
$wrapper | add-member NoteProperty Manufacturer $manufacturer
$wrapper | add-member NoteProperty Model $model
$wrapper | add-member NoteProperty OS $osType
$wrapper | add-member NoteProperty OSArch $osArc
$wrapper | add-member NoteProperty Processor $processor
$wrapper | add-member NoteProperty RAM $ram"GBs"
$wrapper | add-member NoteProperty VideoCard $videoCardName
$wrapper | add-member NoteProperty SerialNumber $serialNumber
$wrapper | add-member NoteProperty RetrievedOn $todaysDate
Export-Csv -InputObject $wrapper -Path $currentDir"\$hostname"Info".csv" -NoTypeInformation -Append

write-host -foreGroundColor green "Finished. $hostname"Info".CSV is located where the script was run from. Press Enter to exit"
read-host

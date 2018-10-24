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
$win7Count = 0
$win8Count = 0
$win10Count = 0
$totalPC = 0

write-host -foregroundColor yellow "Working...please wait"
foreach ($servers in $computersArray){
if(test-connection -ComputerName $servers -Count 1 -quiet)
{
write-host -foreGroundColor green "$servers found, getting OS info"
$os = get-wmiobject Win32_OperatingSystem -computerName $servers
$osType = $os.caption

if($osType -like "*Windows 7*")
{
$win7Count++
$totalPC++
}

elseif($osType -like "*Windows 8*")
{
$win8Count++
$totalPC++
}
elseif($osType -like "*Windows 10*")
{
$win10Count++
$totalPC++
}

}
else
{
write-host -foreGroundColor yellow "$servers is not online"
}
}
[int]$percent7 = ($win7Count / $totalPC) * 100
[int]$percent8 = ($win8Count / $totalPC) * 100
[int]$percent10 = ($win10Count / $totalPC) * 100

write-host -foreGroundColor green "$win7Count Computers running Windows 7 | $percent7%" 
write-host -foreGroundColor green "$win8Count Computers running Windows 8 | $percent8%"
write-host -foreGroundColor green "$win10Count Computers running Windows 10 | $percent10%"
write-host -foreGroundColor green "There are $totalPC computers total"


read-host "Press Enter to Exit"

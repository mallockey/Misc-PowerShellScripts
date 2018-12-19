try{
import-module ActiveDirectory -erroraction stop
}
catch{
write-host -foregroundcolor red "Run from a domain controller"
exit
}
try
{
    $servers = get-adcomputer -filter * -searchbase $ou| select -expandproperty name
}
catch
{
    write-host -foreGroundColor red "OU not correct please verify OU and rerun."
    read-host 
    exit
}
$currentPath = Get-Location
$currentPath = $currentPath.path
$tableName = "DiskDrives"

write-host "Getting disk drive space from computers in computers.txt"

#Create Table object
$table = New-Object system.Data.DataTable “$tableName”

#Define Columns
$col1 = New-Object system.Data.DataColumn ComputerName,([string])
$col2 = New-Object system.Data.DataColumn DriveLetter,([string])
$col3 = New-Object system.Data.DataColumn DriveLabel,([string])
$col4 = New-Object system.Data.DataColumn FreeSpace,([string])
$col5 = New-Object system.Data.DataColumn TotalSpace,([string])
$col6 = New-Object system.Data.DataColumn PercentFree,([string])
$col7 = New-Object system.Data.DataColumn Status,([string])

$table.columns.add($col1)
$table.columns.add($col2)
$table.columns.add($col3)
$table.columns.add($col4)
$table.columns.add($col5)
$table.columns.add($col6)
$table.columns.add($col7)

foreach($server in $servers){
	if(test-connection $server -quiet -count 1){
	try{
	$allDriveInfo = get-wmiobject -class win32_logicaldisk -computerName $server -errorAction stop
	}
	catch{
		write-host -ForegroundColor red "Error getting disk space from $server"
		continue
	}
	write-host -foregroundcolor green "Getting info from: $server"
		foreach($drive in $allDriveInfo){
		
			if($drive.FreeSpace -eq $null){
			continue
			}

			$driveName = $drive.VolumeName
			$freeSpace = [int]($drive.FreeSpace / 1gb)
			$totalSpace = [int]($drive.Size / 1gb)
			$driveLetter = $drive.DeviceID
			[int]$percentFree = ($freeSpace / $totalSpace) * 100
				if($percentFree -lt 10)
				{
				$diskStatus = "LOW"
				}
				else{
				$diskStatus = "OK"
				}
			[string]$freeSpace += " GBs"
			[string]$totalSpace += " GBs"
			[string]$percentFree +="%"
			$row = $table.NewRow()
			$row.DriveLetter = "$driveLetter" 
			$row.DriveLabel = "$driveName" 
			$row.ComputerName = "$server"
			$row.freeSpace = "$freeSpace"
			$row.TotalSpace = "$totalSpace"
			$row.PercentFree = "$Percentfree"
			$row.Status = "$diskStatus"
			#Add the row to the table
			$table.Rows.Add($row)

		}
	}
	else{
	write-host -foregroundcolor yellow $server is not online
	}
}

$Header = @"
<style>
TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: #6495ED;}
TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: yellow;}
</style>
"@

$table | format-table -AutoSize 

write-host "----------------------------"
write-host "1) Export to CSV"
write-host "2) Export to HTML"
write-host "3) Export to both"

write-host "Enter any other key to exit"
write-host "----------------------------"
$choice = read-host "Enter decision"

if($choice -eq 1){
	$table | export-csv $currentPath"\Diskspace.csv" -noTypeInformation
}
elseif($choice -eq 2){
	$table | convertto-HTML -prop ComputerName, DriveLetter, DriveLabel, FreeSpace, TotalSpace, PercentFree, Status -head $Header -Title "Disk Drive Space" | out-file $currentPath"\Diskspace.html"
}
elseif ($choice -eq 3){
	$table | export-csv $currentPath"\Diskspace.csv" -noTypeInformation
	$table | convertto-HTML -prop ComputerName, DriveLetter, DriveLabel, FreeSpace, TotalSpace, PercentFree, Status -head $Header -Title "Disk Drive Space" | out-file $currentPath"\Diskspace.html"
}
else{
	exit
}


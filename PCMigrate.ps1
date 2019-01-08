<#
Author: Josh Melo - joshuamelo1126@gmail.com
Last Updated: 1/8/18
Run this script as the user whose profile is being migrated. This script is used for 
collecting data for a user before migrating domains/PCs. This is used as reassurance to know
what data(if) was missing after the migration.
The following is checked for and exported to a CSV:
1)Username/Domain
2)Default Web Browser
3)Printers(checks for default. Send to PDF/Microsoft/Fax are omitted.
4)Mapped drives. Letters and locations
5)*PST files under the users profile ONLY. The name and path are noted in the table*
6)User's profile path via registry and size of each folder				
#>
#Variable Declarations
$currentLocation = get-location 
$currentLocation = $currentLocation.path
$userDomain = $env:userDomain
$userName = $env:username
$currentUserProfile = $env:USERPROFILE
#Test paths
$testKits = Test-Path -path "C:\Kits"

$table = New-Object system.Data.DataTable “$userInfo”
$col1 = New-Object system.Data.DataColumn Info,([string])
$col2 = New-Object system.Data.DataColumn Value1,([string])
$col3 = New-Object system.Data.DataColumn Value2,([string])
$col4 = New-Object system.Data.DataColumn Value3,([string])

$table.columns.add($col1)
$table.columns.add($col2)
$table.columns.add($col3)
$table.columns.add($col4)

$row = $table.NewRow()
$row.Info ="UserInfo"
$row.Value1=$userDomain
$row.Value2=$userName
$table.rows.Add($row)

function getSizeOfFolder{
    Param(
    $directory
    )
    try{
    $total = get-childitem $directory -recurse | measure-object -Property Length -sum -erroraction stop | select -expandproperty sum
    }
    catch{
    return [String]$total = "N/A"
    }
    if($total -gt 1000000000){  
        $total =  [math]::Round($total / 1gb, 2)
        return [String]$total += "GB"
    }
     elseif($total -gt 1000000){
        $total =  [math]::Round($total / 1mb,2)
        return [String]$total += "MB"
    }
    else{
        return [String]$total = "<1MB"
    }
}
if($testKits -eq $false){
    write-host("-------------------------")
    $createKits = new-item -ItemType directory -Path C:\kits\
}
#Queries Registry value to determine default web browser basedon ProgID value.
function getDefaultBrowser{
    write-host "Getting Default Browser"
    try{
    $defaultBrowser = Get-Itemproperty -path hkcu:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.html\UserChoice\ -ErrorAction stop | select -expandproperty Progid
    }
    catch{
    $defaultBrowser = "Not set"
    return $defaultBrowser
    }
    if($defaultBrowser -like "*Chrome*"){
        $defaultBrowser = "Google Chrome"
    }
    elseif ($defaultBrowser -like "*FireFox*"){
    $defaultBrowser = "Mozilla FireFox"
    }
    elseif($defaultBrowser -like "*APPX*"){
    $defaultBrowser = "Microsoft Edge"
    }
    elseif ($defaultBrowser -like "*IE*" -or $defaultBrowser -like "*HTML*"){
    $defaultBrowser = "Internet Explorer"
    }
    else{
    $defaultBrowser ="Default Browser Unknown"
    }
    return $defaultBrowser
}
#Gets installed printers on PC, excludes Fax, anything containing Microsoft and Send
$printerArray = Get-WmiObject -class win32_printer | where-object {$_.Name -notlike "*Fax*" -and $_.Name -notlike "*Microsoft*" -and $_.Name -notlike "*send*"} 
write-host "Getting Printer info..."
if($printerArray -eq $null){
    write-host "No installed printers."
}
    foreach($printer in $printerArray){
        $row = $table.NewRow()
        $printerName = $printer.Name #DONT KNOW WHY THIS IS NECESSARY BUT OTHERWISE I GET CIM INFO WITHOUT PUTTING IN VARIABLE
        $row.Info ="Installed Printer"
        $row.Value1="$printerName"
            if($printer.default -eq $true){
            $defaultPrinter = $printer.Name
            $row.value2="Default"
            }
        $table.rows.Add($row)
    }
$drives = Get-WmiObject -class win32_mappedlogicaldisk
write-host "Getting Mapped Drives info..."
foreach ($drive in $drives){
    $row = $table.NewRow()
    $row.Info = "Mapped Drive"
    $driveLetter = $drive.Name
    $row.Value1 ="$driveLetter"
    $row.Value2 = $drive.providername
    $table.Rows.Add($row)   
}
#Looks through users profile for .PSTs
$PSTS = Get-ChildItem $currentUserProfile -Recurse -ErrorAction silentlycontinue -Filter '*.pst' 
$totalPSTS = 0
write-host "Getting PST info..."
    if($PSTS -eq $null){
    write-host "No PSTs found under $currentUserProfile"
    }
    else{
        foreach ($pst in $psts){
        $totalPSTS++
        $row = $table.NewRow()
        $row.Info = "PST"
        $PSTName = $PST.name
        $PSTDirectory = $PST.directory
        $row.Info ="PST"
        $row.Value1 = "$PSTName"
        $row.Value2 = "$PSTDirectory"
        $table.Rows.Add($row)   
        }
    }
#Finds users profile folders in registry
$usersFoldersInRegistry = Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
$usersFoldersArray = @()
write-host "Getting user folder info..."
$myDocuments = $usersFoldersInRegistry.Personal
$usersFoldersArray+= $myDocuments
$usersDesktop = $usersFoldersInRegistry.Desktop
$usersFoldersArray+= $usersDesktop
$usersFavorites = $usersFoldersInRegistry.Favorites
$usersFoldersArray +=$usersFavorites
$usersVideos = $usersFoldersInRegistry."My Video"
$usersFoldersArray += $usersVideos
$usersPictures = $usersFoldersInRegistry."My Pictures"
$usersFoldersArray += $usersPictures
$usersMusic = $usersFoldersInRegistry."My Music"
$usersFoldersArray += $usersMusic
    foreach($folder in $usersFoldersArray){  
    $sum = getSizeOfFolder $folder
    $row = $table.NewRow()
    $row.Info ="User Profile Path"
    $row.Value1="$folder"
    $row.value2 = "$sum"
    $table.rows.Add($row)
    }
$currentDefaultBrowser = getDefaultBrowser
write-host "Getting Browser Info..."
$row = $table.NewRow()
$row.Info ="Default Web Browser"
$row.Value1="$currentDefaultBrowser"
$table.rows.Add($row)

write-host "Compling table..."
$table | format-table -AutoSize

write-host "Exporting to: $currentLocation\UserInformation.csv"
$table | export-csv -path $currentLocation"\UserInformation.csv" -notypeinformation

read-host

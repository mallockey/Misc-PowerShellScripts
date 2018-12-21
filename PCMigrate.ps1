
#Create Table object
$currentLocation = get-location 
$currentLocation = $currentLocation.path
$userDomain = $env:userDomain
$userName = $env:username
$currentUserProfile = $env:USERPROFILE
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
$row.Value1="Domain: $userDomain"
$row.Value2="Username: $userName"
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
    elseif($total -gt 1000){
        $total = [math]::Round($total / 1kb,2)
        return [String]$total += "KB"
	}
	else{
		return [String]$total+="Bytes" 
	}
    
}

$testKits = Test-Path -path "C:\Kits"
if($testKits -eq $false){
    write-host("-------------------------")
    $createKits = new-item -ItemType directory -Path C:\kits\
}
#Queries Registry value to determine default web browser.
#Backs up Bookmarks to C:\Kits\ for each browser
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
    write-host "Default Browser is Mozila FireFox"
    $defaultBrowser = "Mozilla FireFox"
    }
    elseif ($defaultBrowser -like "*IE*"){
    write-host "Default Web Browser is Internet Explorer" 
    $defaultBrowser = "Internet Explorer"
    }
    elseif($defaultBrowser -like "*APPX*"){
    write-host "Default Web Browser is Microsoft Edge"
    $defaultBrowser = "Microsoft Edge"
    }
    else{
    write-host "Default Browser Unknown"
    }
    return $defaultBrowser
}

$printerArray = Get-WmiObject -class win32_printer 
write-host "Getting Printer info..."
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

$PSTS = Get-ChildItem $currentUserProfile -Recurse -ErrorAction silentlycontinue -Filter '*.pst' 
$totalPSTS = 0
write-host "Getting PST info..."
    if($PSTS -eq $null){
    write-host "No PSTs found"
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

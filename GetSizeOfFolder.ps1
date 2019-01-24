Param(
  [Parameter(Mandatory=$true)]
  $PathToCheck
)
$testPath = test-path $PathToCheck
if($PathToCheck -eq $null -or $testPath -eq $false){
    write-host -ForegroundColor red "$pathTocheck was not valid, please rerun script."
    exit
}

$currentPath = get-location | select -expandproperty path
$table = New-Object system.Data.DataTable “$tableName”
$col1 = New-Object system.Data.DataColumn Folder,([string])
$col2 = New-Object system.Data.DataColumn Space,([string])
$table.columns.add($col1)
$table.columns.add($col2)

function getSizeOfFolder{
    Param(
    $directory
    )
    try{
    $total = get-childitem $directory -recurse -errorAction SilentlyContinue | measure-object -Property Length -sum -erroraction stop | select -expandproperty sum
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

$usersFoldersArray = get-childitem -name $pathToCheck

foreach($folder in $usersFoldersArray){  
	$testFolder = test-path -path $pathToCheck\$folder -pathType container
	if($testFolder -eq $true){
	Write-Progress -Activity "Getting Size of $pathToCheck" -CurrentOperation "Current Folder: $folder"
	$sum = getSizeOfFolder $pathToCheck\$folder
	
	$row = $table.NewRow()
	$row.Folder = "$folder" 
	$row.Space = "$sum" 
	$table.Rows.Add($row)
	}
	else{
	continue
	}

}

$table | format-table -AutoSize 
write-host -foregroundcolor yellow "INFO: Folders that the user running the script from, doesn't have access to, will not be accounted for."
write-host -foregroundcolor yellow "INFO: CSV was exported to $currentPath"
$table | export-csv "$currentPath\UsersFoldersSizes.csv" -noTypeInformation
read-host

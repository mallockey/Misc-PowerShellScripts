Param(
  $SizeInUnits = "gb",
  $BackUpPath = "",
  $NoAppData
)
$sizeInUnits = $sizeInUnits.ToUpper()
$sizeHeader = "SizeIn" + $sizeInUnits
$sizeInUnits = "1" + $sizeInUnits
$computerArray = Get-Content "C:\kits\Workstations.txt"
$arrayOfPCs = [System.Collections.ArrayList]@()
function createArrayOfObjects{
  Param(
    $computer,
    $userFolderName,
    $totalPerUser
  )
  $tempObj = New-Object -TypeName PSObject 
  $tempObj | Add-Member NoteProperty -Name "PCName" -Value $computer
  $tempObj | Add-Member NoteProperty -Name "User" -Value $userFolderName
  $tempObj | Add-Member NoteProperty -Name "$sizeHeader" -Value $totalPerUser
  $arrayOfPCs.Add($tempObj) | out-null 
  return $totalPerUser
}
function addTotalPerPC{
  Param(
    $computer,
    $totalPerPC
  )
  $tempObj = New-Object -TypeName PSObject 
  $tempObj | Add-Member NoteProperty -Name "PCName" -Value "TotalForPC:"
  $tempObj | Add-Member NoteProperty -Name "User" -Value $computer
  $tempObj | Add-Member NoteProperty -Name "$sizeHeader" -Value $totalPerPC
  $arrayOfPCs.Add($tempObj) | out-null
  return $totalPerPC
}
foreach($computer in $computerArray){
  if(Test-Connection -ComputerName $computer -quiet -count 1){
    $currentPCPath = New-Item -Path "$BackUpPath\$computer" -ItemType Directory
    $totalPerPC = 0
    $userFoldersArray = Get-ChildItem "\\$computer\C$\users\" | select -expandproperty fullname
    foreach($userFolder in $userFoldersArray){
      $totalPerUser = get-childitem $userFolder -recurse -ErrorAction SilentlyContinue | measure-object -Property Length -sum -erroraction stop | select -expandproperty sum 
      $totalPerUser =  [math]::Round($totalPerUser / $sizeInUnits, 2)
      $userFolderName = Split-Path -Leaf $userFolder
      $totalPerPC += createArrayOfObjects -computer $computer -userFolderName $userFolderName -totalPerUser $totalPerUser
      $currentUserPath = Copy-Item -Path $userFolder -Exclude "$userFolder\appdata" -Destination $currentPCPath\$userFolderName -Recurse
    }
  $totalTotalTOTAL += addTotalPerPC -totalPerPC $totalPerPC -computer $computer
  }
}
$arrayOfPCs | ft

$totalFreeDiskSpace = Get-WmiObject -Class win32_volume | Where-Object {($_.Name -eq ($BackUpPath.SubString(0,3)))}| select -ExpandProperty FreeSpace

$totalFreeDiskSpace = [math]::Round($totalFreeDiskSpace / $sizeInUnits, 2)
$SizeInUnits = $SizeInUnits.Replace("1","")
if($totalTotalTOTAL -gt $totalFreeDiskSpace){
  Write-Warning "There is not enough disk space to copy all of users folders over."
  Write-Warning "Either change to a different drive or clean up disk space before copying profiles"
  Write-Warning "You can also try using the -NoAppdata parameter to exclude appdata folder for"
}

write-host "Total Amount of Disk Space needed: $totalTotalTOTAL$SizeInUnits"
Write-Host "Total Amount of Disk Space available: $totalFreeDiskSpace$SizeInUnits"
  



      
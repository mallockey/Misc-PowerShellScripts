Param(
  $SizeInUnits = "gb",
  $OnlyLoggedOnUser = $false
)
$sizeInUnits = $sizeInUnits.ToUpper()
$sizeHeader = "SizeIn" + $sizeInUnits
$sizeInUnits = "1" + $sizeInUnits
$computerArray = Get-Content "C:\kits\Workstations.txt"
$arrayOfPCs = [System.Collections.ArrayList]@()
$userFoldersArray = Get-ChildItem "C:\users\" | select -expandproperty fullname
foreach($computer in $computerArray){
  $totalPerComputer = 0
  $totalPerUser = 0
  if($OnlyLoggedOnUser -eq $true){
    
    try{
      $loggedOnUser = Get-WmiObject -class win32_computersystem -ComputerName $computer | select -ExpandProperty username
    }
    catch{
      $loggedOnUser = "N/A"
    }
    $tempIndex = $loggedOnUser.IndexOf("\") + 1
    $loggedOnUser = $loggedOnUser.SubString($tempIndex)
    $userFoldersArray = Get-ChildItem "C:\users\"$loggedOnUser | select -ExpandProperty fullname
    $totalPerUser = get-childitem $userFoldersArray -recurse -errorAction SilentlyContinue | measure-object -Property Length -sum -erroraction stop | select -expandproperty sum
    $totalPerUser =  [math]::Round($totalPerUser / $sizeInUnits, 2)
    write-host $totalPerUser
    $tempObj = New-Object -TypeName PSObject 
    $tempObj | Add-Member NoteProperty -Name "PCName" -Value $computer
    $tempObj | Add-Member NoteProperty -Name "User" -Value $userFoldersArray
    $tempObj | Add-Member NoteProperty -Name "$sizeHeader" -Value $totalPerUser
    $arrayOfPCs.Add($tempObj) | out-null  
  }
  else{
    foreach($userFolder in $userFoldersArray){
        $totalPerUser = get-childitem $userFolder -recurse -errorAction SilentlyContinue | measure-object -Property Length -sum -erroraction stop | select -expandproperty sum
        $totalPerUser =  [math]::Round($totalPerUser / $sizeInUnits, 2)
        $tempObj = New-Object -TypeName PSObject 
        $tempObj | Add-Member NoteProperty -Name "PCName" -Value $computer
        $tempObj | Add-Member NoteProperty -Name "User" -Value $userFolder
        $tempObj | Add-Member NoteProperty -Name "$sizeHeader" -Value $totalPerUser
        $arrayOfPCs.Add($tempObj) | out-null   
    }
    $totalPerComputer+=$totalPerUser
  }
}
$arrayOfPCs |  ft


      
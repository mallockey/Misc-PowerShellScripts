Param(
  [Parameter(ParameterSetName = "AD")][ValidateNotNullOrEmpty()][String]$SpecifyOU,
  [Parameter(ParameterSetName = "AD")][Switch]$WorkstationsOnly,
  [Parameter(ParameterSetName = "AD")][Switch]$ServersOnly,
  [Parameter(ParameterSetName = "InputFile")][ValidateNotNullOrEmpty()][String]$InputFile,
  [parameter(ParameterSetName = "Computers",ValueFromPipeline)][Array]$ComputerName,
  [String]$OutputFile,
  [Switch]$UseCIM
)

$ErrorActionPreference = 'Stop'

if($SpecifyOU -or $WorkstationsOnly -or $ServersOnly){
  try{
    Import-Module ActiveDirectory
  }catch{
    Write-Warning "Unable to import the ActiveDirectory Module"
    Write-Warning "Make sure you are running this from a domain controller"
    exit
  }
  if($SpecifyOU){
    try{
      $allComputers = (Get-ADComputer -Filter * -SearchBase $SpecifyOU).Name
    }
    catch{
      Write-Warning "OU not correct please verify OU and rerun."
      exit
    }
  }elseif($WorkstationsOnly){
    $allComputers = (Get-ADComputer -Filter {OperatingSystem -NotLike '*server*' -and Enabled -eq $True}).Name
  }elseif($ServersOnly){
    $allComputers = (Get-ADComputer -Filter {OperatingSystem -Like '*server*' -and Enabled -eq $True}).Name
  }else{
    $allComputers = (Get-ADComputer -Filter {Enabled -eq $True}).Name
  }
}elseif($inputFile){
  try{
    $allComputers = Get-Content $InputFile 
  }catch{
    Write-Warning "$inputFile is not a valid list of workstations."
  }
}elseif($computerName){
  $allComputers = $ComputerName
}

  $resultsArray = @()

  $objProp = [Ordered]@{
    ComputerName = $null
    DriveLetter = $null
    DriveLabel = $null
    FreeSpace = $null
    TotalSpace = $null
    PercentFree = $null
    Status = $null
    Online = $null
  }

function get-DiskStats {
  Param(
    $computerName
  )

  if($UseCIM){
    $cmdLetToTry = "Get-CimInstance -ClassName Win32_LogicalDisk -ComputerName $computerName"
  }else{
    $cmdLetToTry = "Get-WMIObject -Class Win32_LogicalDisk -ComputerName $computerName"
  }
  try{
    $allDriveInfo = Invoke-Expression -Command $cmdLetToTry
  }catch{
    $allDriveInfo = $null
  }

  return $allDriveInfo
}

for($i=0; $i -lt $allComputers.length; $i++){

  [Int]$currentPercent = ($i / $allComputers.length) * 100
  Write-Progress -Activity "Getting disk info from $($allComputers[$i])" -CurrentOperation "$currentPercent% completed"
  
  $computerObj = New-Object -TypeName PSObject -Prop $objProp
  $currentComputer = $allComputers[$i]
  $computerObj.ComputerName = $currentComputer

  if(Test-Connection $currentComputer -Quiet -Count 1){
    
    $computerObj.Online = "Online"
    $allDriveInfo = get-DiskStats -computerName $currentComputer

    if($allDriveInfo){
      foreach($drive in $allDriveInfo){
        if($drive.DriveType -ne 3){
          continue
        }
    
        $computerObj = New-Object -TypeName PSObject -Prop $objProp
      
        $freeSpace = ([int]($drive.FreeSpace / 1gb))
        $freeSpaceString = $freeSpace.ToString() + "GBs"
        
        $totalSpace = ([int]($drive.Size / 1gb))
        $totalSpaceString = $totalSpace.ToString() + "GBs"
        
        $percentFree = ([Int](($freeSpace / $totalSpace) * 100))
        $percentFreeString = $percentFree.ToString() + "%"

        $computerObj.FreeSpace = $freeSpaceString     
        $computerObj.ComputerName = $currentComputer
        $computerObj.Online = "Online"
        $computerObj.DriveLetter = $drive.DeviceID
        $computerObj.DriveLabel = $drive.VolumeName
        $computerObj.TotalSpace = $totalSpaceString
        $computerObj.PercentFree = $percentFreeString
        $computerObj.Status = "OK"

        if ($percentFree -lt 10){
          $computerObj.Status = "LOW"
        }

        $resultsArray += $computerObj
      }
    }else{

      $computerObj.Online = "Online"

      if($UseCIM){
        $computerObj.Status = "Unable to get disk info via CIM"
      }else{
        $computerObj.Status = "Unable to get disk info via WMI"
      }
      $computerObj.Online = "Online"
      $resultsArray += $computerObj

    }
  }else{

    $computerObj.Online = "Offline"
    $resultsArray += $computerObj
  }
}

$resultsArray | Sort-Object Online -Descending | Format-Table

if($OutputFile){
  $checkIfCSV = $OutputFile.Substring($OutputFile.Length -3)
  if($checkIfCSV -ne "csv"){
    $OutputFile = $OutputFile.Replace($checkIfCSV,"csv")
  }
  $resultsArray | Sort-Object Online -Descending | Export-Csv $outputFile -NoTypeInformation
}

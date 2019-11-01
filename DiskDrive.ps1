param(
  [Switch]$ActiveDirectory,
  [String]$WorkstationsOU,
  [String]$InputFile,
  [String]$OutputFile,
  [Switch]$WinRM,
  [Switch]$AllComputers,
  [Switch]$WorkstationsOnly,
  [Switch]$ServersOnly
)
$ErrorActionPreference = 'Stop'

if($ActiveDirectory){
  try{
    Import-Module ActiveDirectory -erroraction stop
  }catch{
    Write-Output "Run from a domain controller"
    exit
  }
  if($AllComputers){
    $allComputers = Get-ADComputer -Filter {Enabled -eq $True} | Select-Object -expandproperty Name
  }elseif($WorkstationsOU){
    try{
      $allComputers = Get-ADComputer -Filter {Enabled -eq $True} -SearchBase $WorkstationsOU | Select-Object -expandproperty Name
    }
    catch{
      Write-Output "OU not correct please verify OU and rerun."
      exit
    }
  }elseif($WorkstationsOnly){
    $allComputers = Get-ADComputer -Filter {OperatingSystem -NotLike '*server*' -and Enabled -eq $True} | Select-Object -ExpandProperty Name
  }elseif($ServersOnly){
    $allComputers = Get-ADComputer -Filter {OperatingSystem -Like '*server*' -and Enabled -eq $True} | Select-Object -ExpandProperty Name
  }
}else{
  try{
    $allComputers = Get-Content $InputFile 
  }catch{
    Write-Output "$inputFile is not a valid list of workstations."
  }
}
  
  $currentPath = Get-Location
  $currentPath = $currentPath.path
  $resultsArray = @()
  
  $objProp = @{
    ComputerName = $null
    DriveLetter = $null
    DriveLabel = $null
    FreeSpace = $null
    TotalSpace = $null
    PercentFree = $null
    Status = $null
    Online = $null
  }

for($i=0; $i -lt $allComputers.length; $i++){
  [Int]$currentPercent = ($i / $allComputers.length) * 100
  Write-Progress -Activity "Getting disk info from $($allComputers[$i])" -CurrentOperation "$currentPercent% completed"
  
  $computerObj = New-Object -TypeName PSObject -Prop $objProp
  $currentComputer = $allComputers[$i]
  $computerObj.ComputerName = $currentComputer

  if(Test-Connection $currentComputer -Quiet -Count 1){

    $computerObj.Online = "Online"
    if($WinRM){
      $allDriveInfo = Get-CimInstance -ClassName Win32_LogicalDisk -ComputerName $currentComputer
    }else{
      $allDriveInfo = Get-WmiObject -Class Win32_LogicalDisk -ComputerName $currentComputer
    }
    foreach($drive in $allDriveInfo){
      if($drive.FreeSpace -eq $null){
        continue
      }
      $computerObj.DriveLetter = $drive.DeviceID
      $computerObj.DriveLabel = $drive.DriveLabel
      $computerObj.FreeSpace = [int]($drive.FreeSpace / 1gb)
      $computerObj.TotalSpace = [int]($drive.Size / 1gb)
      $computerObj.PercentFree = [Int](($computerObj.FreeSpace / $computerObj.TotalSpace) * 100)
      $computerObj.Status = "OK"
        if($obj.PercentFree -lt 10){
          $computerObj.Status = "LOW"
        }
      }
    }else{
      $computerObj.Online = "Offline"
    }
    $resultsArray += $computerObj
}
$resultsArray | ft
$resultsArray | Export-Csv "$outputFile\DiskDriveInfo.csv" -NoTypeInformation

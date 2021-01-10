param(
  [CmdletBinding()]
  [String]$InstallPath = "C:\Kits",
  [Switch]$FetchLatestAppNames, #Invokes web request to scrape data from checkboxes
  [Array]$SpecifyApps #Specify app names from command line
)

$ErrorActionPreference = "Stop"

$CurrentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if(!($CurrentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))){
    Write-Warning "Script is not running as Administrator"
    Write-Warning "Please rerun this script as Administrator."
    exit
}
function Get-LatestNiniteAppsLists {

  $NinitePageRaw = Invoke-RestMethod -Uri "ninite.com"
  $ParsedNinitePage = $NinitePageRaw -split '\r?\n'
  $ParsedNinitePage = $ParsedNinitePage | Where-Object {$_ -like "*name=`"apps`"*"}
  
  $AppsToInstall = $ParsedNinitePage | ForEach-Object {
    $IndexOfValue = ($_.IndexOf("value=")) + 7
    $CurrentApp = $_.SubString($IndexOfValue)
    $CurrentApp.Remove($CurrentApp.Length -1)
  }

  $AppsToInstall | Out-File ".\ListOfApps.txt"
  
}

if($FetchLatestAppNames){ 
  Get-LatestNiniteAppsLists
}

if(!($SpecifyApps)){

  if(!(Test-Path ".\ListOfApps.txt")){
    Write-Warning "ListOfApps.txt missing in directory"
    Write-Warning "You can regenerate it by using the -FetchLatestAppNames parameter"
    exit
  }

  $ListOfApps = Get-Content ".\ListOfApps.txt"
  $ListOfApps | Sort-Object
  Write-Output "========================"

  $UsersApps = @()
  while($UserInput -ne "done"){
    $UserInput = Read-Host -Prompt "Enter the apps you'd like to install from the list above one at a time, enter done to finish"
    if($UserInput -eq "done"){
      break
    }
    if($ListOfApps -notcontains $UserInput){
      Write-Warning "$UserInput was not listed in the list above"
      continue
    }
    $UsersApps += $UserInput
  }
}

if(!($SpecifyApps)){
  $UsersApps | ForEach-Object {
    $InstallString += $_ + "-"
  }
}else{
  $SpecifyApps | ForEach-Object {
    $InstallString += $_ + "-"
  }
}

$InstallString = $InstallString.Remove($InstallString.Length -1) #Remove the last -

try{
  if(!(Test-Path -Path $InstallPath)){
    New-Item -ItemType Directory -Path $InstallPath
  }
  Invoke-RestMethod -Uri "https://ninite.com/$($InstallString)/ninite.exe" -OutFile "$InstallPath\Ninite.exe"
  Start-Process "$InstallPath\Ninite.exe" 
}catch{
  Write-Warning "There was an error downloading/installing from ninite:"
  Write-Warning $_
}

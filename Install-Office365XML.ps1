Param(
  [Parameter(Mandatory=$True)]
  [ValidateNotNull()]
  [String]$ConfiguratonXMLFile,
  [String]$OfficeInstallerPath,
  [Switch]$Silent
)

$CurrentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
If(!($CurrentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))){
    Write-Warning "Script is not running as Administrator"
    Write-Warning "Please rerun this script as Administrator."
    Exit
}

If(!($Silent)){
  $VerbosePreference = "Continue"
}Else{
  $WarningPreference = "SilentlyContinue"
}

If(!(Test-Path $ConfiguratonXMLFile)){
  Write-Warning "The configuration XML file is not a valid file"
  Write-Warning "Please check the path and try again"
  Exit
}

$ErrorActionPreference = "Stop"
$ODTInstallLink = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_12130-20272.exe"
$OfficeInstallerPath = "C:\Scripts\OfficeInstall"

If(-Not(Test-Path $OfficeInstallerPath)){
  New-Item -Path $OfficeInstallerPath -ItemType Directory -ErrorAction Stop | Out-Null
}

#Download the Office Deployment Tool
Write-Verbose "Downloading the Office Deployment Tool..."
Try{
  Invoke-WebRequest -Uri $ODTInstallLink -OutFile "$OfficeInstallerPath\ODTSetup.exe"
}Catch{
  Write-Warning "There was an error downloading the Office Deployment Tool."
  Write-Warning "Please verify the below link is valid:"
  Write-Warning $ODTInstallLink
  Exit
}

#Run the Office Deployment Tool
Try{
  Write-Verbose "Running the Office Deployment Tool..."
  Start-Process "$OfficeInstallerPath\ODTSetup.exe" -ArgumentList "/quiet /extract:$OfficeInstallerPath" -Wait
}Catch{
  Write-Warning "Error running the Office Deployment Tool. The error is below:"
  Write-Warning $_
}

#Run the install
Try{
  Write-Verbose "Downloading and installing Office 365"
  $OfficeInstall = Start-Process "$OfficeInstallerPath\Setup.exe" -ArgumentList "/configure $ConfiguratonXMLFile" -Wait -PassThru
}Catch{
  Write-Warning "Error running the Office install. The error is below:"
  Write-Warning $_
}

If($OfficeInstall.ExitCode -ne 0){
  Write-Warning "Office install may have installed with errors."
}Else{
  Write-Verbose "Office $($OfficeVersionEdition) installed successfully"
}

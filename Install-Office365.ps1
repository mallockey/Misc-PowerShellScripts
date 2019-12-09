Param(
  [Switch]$OfficePP32Bit,
  [Switch]$OfficeVersion
)

$CurrentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
If(!($CurrentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))){
    Write-Warning "Script is not running as Administrator"
    Write-Warning "Please rerun this script as Administrator."
    Exit
}

$ErrorActionPreference = "Stop"
$ODTInstallLink = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_12130-20272.exe"
$OfficeInstallerPath = "C:\Scripts\OfficeInstall"

If(-Not(Test-Path $OfficeInstallerPath)){
  New-Item -Path $OfficeInstallerPath -ItemType Directory -ErrorAction Stop | Out-Null
}

#Download the Office Deployment Tool
Write-Output "Downloading the Office Deployment Tool..."
Try{
  Invoke-WebRequest -Uri $ODTInstallLink -OutFile "$OfficeInstallerPath\ODTSetup.exe"
}Catch{
  Write-Warning "There was an error downloading the Office Deployment Tool."
  Write-Warning "Please verify the below link is valid:"
  Write-Warning $ODTInstallLink
  Exit
}

Try{
  Write-Output "Running the Office Deployment Tool..."
  Start-Process "$OfficeInstallerPath\ODTSetup.exe" -ArgumentList "/quiet /extract:$OfficeInstallerPath" -Wait
}Catch{
  Write-Warning "Error running the Office Deployment Tool. The error is below:"
  Write-Warning $_
}

[String]$OfficeArch = "64"
$OfficeVersionEdition = "O365ProPlusRetail"

If($OfficePP32Bit){
  $OfficeArch = "32"  
}
If($OfficeVersion){
  $OfficeVersionEdition = "O365BusinessRetail"
}

$OfficeXML = [XML]@"
<Configuration>
  <Add OfficeClientEdition="$OfficeArch">
    <Product ID="$OfficeVersionEdition">
      <Language ID="en-us" />
    </Product>
  </Add>  
  <Display Level="None" AcceptEULA="FALSE" />
</Configuration>
"@

#Save the XML file
$OfficeXML.Save("$OfficeInstallerPath\OfficeInstall.xml")
#Run the install
Try{
  Write-Output "Downloading and installing Office 365"
  $OfficeInstall = Start-Process "$OfficeInstallerPath\Setup.exe" -ArgumentList "/configure $OfficeInstallerPath\OfficeInstall.xml" -Wait -PassThru
}Catch{
  Write-Warning "Error running the Office install. The error is below:"
  Write-Warning $_
}

If($OfficeInstall.ExitCode -ne 0){
  Write-Warning "Office install may have installed with errors."
}Else{
  Write-Output "Office $($OfficeVersionEdition) installed successfully"
}

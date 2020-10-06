<#
    Installs .NET 3.5 SP1 from DISM
#>

$ErrorActionPreference = "Stop"

$DotNetInstall = Start-Process "DISM" -ArgumentList "/Online /Enable-Feature /FeatureName:NetFx3 /All /NoRestart" -Wait -PassThru

if(Get-Childitem -Path "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP" | Where-Object -FilterScript {$_.name -match "v3.5"}){
	$True
}else{
	$False
}

#Author: Josh Melo 6/7/18
#Run this script as administrator.

function successText($output)
	{
	$output = write-host -ForeGroundColor green "SUCCESS:"$output
	return $output
	}

function failText($output)
	{
	$output = write-host -ForeGroundColor red "ERROR:"$output
	return $output
	}
	
function infoText($output)
	{
	$output = write-host -foreGroundColor yellow "INFO:"$output
	return $output
	}
	

write-host "===========================================New PC Assistant============================================================"
write-host -foreGroundColor yellow("You should run this script AFTER the PC has been joined to the domain and from an account that has")
write-host -foreGroundColor yellow("access to the remote kits folder")
write-host "======================================================================================================================="

[int]$numPrograms = 0
$officeConfigLink = "$psscriptroot\config.xml"
write-host -foreGroundColor green ("This script will:")
write-host -foreGroundColor green ("1)Install software")
write-host "======================================================================================================================="
infoText("Please make sure all install files are in the folder being copied.")
$decision = read-host "Do you want to copy kits folder from a server(Use y or n)"
if($decision -eq 'y')

{
	write-host "======================================================================================================================="
	
	#Copy Block
	$remoteKitsPath = read-host "Enter the path where the kits folder is without quotes" -ErrorAction SilentlyContinue
		try{
		$testRemoteKitsPath = test-path -path $remoteKitsPath -ErrorAction SilentlyContinue
		}
		catch{
		failText("No input was entered, please rerun")
		exit
		}
	if ($testRemoteKitsPath -eq $true -And [string]::IsNullOrWhiteSpace($remoteKitsPath) -eq $false -And $remoteKitsPath -ne $null) {
		successText("Remote kits folder found.")
		write-host "================================================Directory of $remoteKitsPath==========================================="
		get-childitem $remoteKitsPath -name
		write-host "======================================================================================================================="
		$confirmation = Read-Host "Are you sure you this is the correct path? ""$remoteKitsPath"" (Use y for yes)"
		$remoteKitsLeaf = split-path $remoteKitsPath -leaf
		$localKitsPath = "C:\kits\"
		$testfullLinkLocal = test-path $localKitsPath$remoteKitsleaf
		
		if ($confirmation -eq 'y') { 
			
		 
			if ($testfullLinkLocal -eq $true) {
				failText ("Folder already exists on in C:\Kits, please rename and rerun. ")
				exit
				}
							for ($a=1; $a -lt 100; $a++) {
							Write-Progress -Activity "Copying $remoteKitsLeaf to C:\Kits"  -PercentComplete $a -CurrentOperation "$a% complete"
					start-sleep -Milliseconds 25
					
					
				 
						}
						copy-item -path $remoteKitsPath -destination "c:\kits" -recurse 
						Write-Progress -Activity "Copying $remoteKitsLeaf to C:\kits Completed." -Completed
			
		} else {
			failText("Please verify remote kits folder and rerun script")
			exit
		}
	} else {
		failText("Path is not correct. Please verify path and rerun script")
		read-host "Press enter to exit"
		exit
	}
	
#End of Copy Block

#Array Instantiation 
######################################################Remote Kits Folder Install#############################################################
successText("Copy Complete!")
	try{
	$numPrograms = read-host "Please enter the number of programs you want to install"
	}
	
	catch{
	failText("You did not enter a number, please rerun script")
	exit
	
	}

	$office = get-childitem -path $localkitspath$remoteKitsleaf *office* -name -errorAction silentlyContinue
	if($office -eq "office")
	{
	infoText("Office folder already exists, won't rename.")
	}
	elseif([string]::IsNullOrWhiteSpace($office))
	{
	infoText("No Office setup file found.")
	}
	else{
	rename-item -path "$localkitspath$remoteKitsleaf\$office" -newname Office
	infoText("Office setup file found, renaming folder to just Office for ease of use.")
	}
	
	if($numPrograms -is [int]){
		
		$arrayOfProgramLinks = New-Object -TypeName 'object[]' -ArgumentList $numPrograms
		
		write-host "=========================================Directory of C:\Kits\$remoteKitsLeaf=========================================="
		infoText( "Enter only the setup files unless the setup file is in a subdirectory, do not include quotes. Example:itunes.exe")
		get-childitem $localKitsPath$remoteKitsLeaf -name
		write-host "======================================================================================================================="
		
		for($i=0; $i -lt $numPrograms; $i++)
			{
			$nameOfLink = read-host "Enter the link of the setup file in the kits folder"
			
			$nameOfLink = "\" + $nameOfLink
			$testInstall = test-path -path $localKitsPath$remoteKitsLeaf$nameOfLink
			
				if($testInstall -eq $true){
				$arrayOfProgramLinks[$i] = $localKitsPath + $remoteKitsLeaf + $nameOfLink
			
				}
				else{
				failText("Link is not correct.")
				
				}
			
			}
							
	}
	else{
	write-host "You didn't enter a number"
	}
}
####################################################Local Kits Install###########################################################

	elseif($decision -eq 'n')
	{
		
	$localKitsFolder = read-host "Please enter where the local kits folder is without quotes"
	try{
		$testLocalKitsFolder = test-path -path $localKitsFolder -errorAction silentlycontinue
		}
	catch{
	failText("No input was entered, please rerun")
	exit
	}


		if($testLocalKitsFolder -eq $true)
		{
		successText("Local folder found.")
		write-host "======================================================================================================================="
		get-childitem -path $localKitsFolder -name
		write-host "======================================================================================================================="
		$confirmation = Read-Host "Are you sure you this is the correct path? ""$localKitsfolder"" (Use y for yes)"
				
				
			if ($confirmation -eq 'y') { 
				
				$office = get-childitem -path $localkitsfolder *office* -name -errorAction silentlyContinue
				if($office -eq "office")
				{
				infotext("Office folder already exists, won't rename.")
				}
				elseif([string]::IsNullOrWhiteSpace($office))
				{
				infoText("No Office setup file found.")
				}
				else{
				rename-item -path "$localKitsFolder\$office" -newname Office
				infoText( "Office setup file found, renaming folder to just Office for ease of use.")
			
				}
				}
				
			else{
			failText("Please confirm local path and rerun script")
			exit
			
			}

				try{
				$numPrograms = read-host "Please enter the number of programs you want to install"
				}
				
				catch{
				failText("You did not enter a number, please rerun script")
				exit
				
				}

				if($numPrograms -is [int])
				{
						
					$arrayOfProgramLinks = New-Object -TypeName 'object[]' -ArgumentList $numPrograms
					
					write-host "============================================Directory of $localKitsFolder=============================================="
					infoText("Enter only the setup files unless the setup file is in a subdirectory, do not include quotes. Example:itunes.exe")
					get-childitem $localKitsFolder -name
					write-host "======================================================================================================================="
					for($i=0; $i -lt $numPrograms; $i++)
					{
						$nameOfLink = read-host "Enter the link of the setup file in the folder"
						$nameOfLink = "\" + $nameOfLink
						
						$testInstall = test-path -path $localKitsFolder$nameOfLink
						
							if($testInstall -eq $true)
							{
							$arrayOfProgramLinks[$i] = $localKitsFolder + $nameOfLink
							}
							else
							{
							failText("Link is not correct")
							}			
					}							
				}

				else
				{
				failText("You did not a number, please rerun script")
				exit
				}
		}
		else{
		failText("Path is incorrect")
		exit
		
		}

	}

else
{
failText("You did not enter y or n, please rerun script")
exit
}
#End Of Array Instantiation
	
#Installation Block
	
	write-host "=================================================Install Block========================================================="
	successText( "Disabling UAC for the time being...")
   	Set-ItemProperty -Path REGISTRY::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System -Name ConsentPromptBehaviorAdmin -Value 0 -Erroraction SilentlyContinue
	
	for($i=0; $i -lt $numPrograms; $i++){
		#MSI 
		
		if($arrayOfProgramLinks[$i] -match '.msi'){
		for ($a=1; $a -lt 100; $a++) {
					Write-Progress -Activity "Installing "$arrayOfProgramLinks[$i] -Id 1  -PercentComplete $a -CurrentOperation "$a% complete"
					start-sleep -Milliseconds 25
						}
		$setupMSI = Start-Process $arrayOfProgramLinks[$i] -ArgumentList "/quiet" -wait -PassThru 
		Write-Progress -Activity "Completed" -Completed -Id 1
		
			if($setupMSI.ExitCode -eq 0)
				{
				successText($arrayOfProgramLinks[$i] + " installed successfully!")
				}
		
			else{
				failText ($arrayOfProgramLinks[$i] +  " did not install successfully, Error Code is: $($setupMSI.ExitCode)")
				}
		
		}
		#END OF MSI
		#NINITE
		elseif($arrayOfProgramLinks[$i] -match 'ninite'){
		$setupNinite = Start-Process $arrayOfProgramLinks[$i] -wait -PassThru
		
		
		if($setupNinite.ExitCode -eq 0)
				{
				successText($arrayOfProgramLinks[$i] + " installed successfully!")
				}
		
		
			else{
				failText ("$arrayOfProgramLinks[$i] did not install successfully, Error Code is: $($setupNinite.ExitCode)")
			}
		}
		#END OF NINITE
		
		#OFFICE
		elseif($arrayOfProgramLinks[$i] -match 'office'){
		$officeConfig = split-path $arrayOfProgramLinks[$i]
		copy-item -path "$psscriptroot\config.xml" -destination $officeConfig
	
		for ($a=1; $a -lt 100; $a++) {
					Write-Progress -Activity "Installing Microsoft Office" -Id 2 -PercentComplete $a -CurrentOperation "$a% complete"
					start-sleep -Milliseconds 25
						}
		
		$setupOffice = Start-Process $arrayOfProgramLinks[$i] -ArgumentList "/config config.xml" -wait -PassThru
			Write-Progress -Activity "Installing Microsoft Office" -Id 2 -Completed
		
		if($setupOffice.ExitCode -eq 0)
				{
				successText($arrayOfProgramLinks[$i] + " installed successfully!")
				}
		
		
			else{
				failText ("$arrayOfProgramLinks[$i] did not install successfully, Error Code is: $($setupOffice.ExitCode)")
			}
		}
		#END OF OFFICE
		
		#BLOOMBERG
		elseif($arrayOfProgramLinks[$i] -match 'sotrt*'){
		for ($a=1; $a -lt 100; $a++) {
					Write-Progress -Activity "Installing Bloomberg..." -Id 3 -PercentComplete $a -CurrentOperation "$a% complete"
					start-sleep -Milliseconds 25
						}	
		$setupBloomberg = start-process $arrayOfProgramLinks[$i] -ArgumentList "/s" -Wait -PassThru
		Write-Progress -Activity "Installing Bloomberg" -Id 3 -Completed
			if($setupBloomberg.ExitCode -eq 0)
				{
				successText($arrayOfProgramLinks[$i] + " installed successfully!")
				}
			else{
				failText ($arrayOfProgramLinks[$i] + " did not install successfully, Error Code is: $($setupEXE.ExitCode)")
			}
		}
		
		#EXE
		elseif($arrayOfProgramLinks[$i] -match '.exe'){
		for ($a=1; $a -lt 100; $a++) {
					Write-Progress -Activity "Installing "$arrayOfProgramLinks[$i]  -Id 4 -PercentComplete $a -CurrentOperation "$a% complete"
					start-sleep -Milliseconds 25
						}	
		$setupEXE = start-process $arrayOfProgramLinks[$i] -ArgumentList "/silent" -Wait -PassThru
		Write-Progress -Activity "Completed" -Id 4 -Completed
			if($setupEXE.ExitCode -eq 0)
				{
				successText($arrayOfProgramLinks[$i] + " installed successfully!")
				}
			else{
				failText ($arrayOfProgramLinks[$i] + " did not install successfully, Error Code is: $($setupEXE.ExitCode)")
			}
		}
		
		#END OF EXE
		
	}
successText("Reenabling UAC")
Set-ItemProperty -Path REGISTRY::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System -Name ConsentPromptBehaviorAdmin -Value 5 -Erroraction SilentlyContinue
 
write-host "======================================================================================================================="
#End Installation Block
write-host "======================================================================================================================="
successText("Disabling users to add or login with Microsoft Accounts")
Set-ItemProperty -Path REGISTRY::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System -Name NoConnectedUser -Value 3 -Erroraction SilentlyContinue

$driveConfirmation = read-host "Do you want to change CD-ROM drive letter?(Use y for yes)"
	if($driveConfirmation -eq "y")
		{
		$driveLetter = read-host "Please enter the current CD-ROM drive letter, it will change to R(Ex:"D:")"
		$doubleConf = read-host "Are you sure you want to change the CD-ROM drive letter?(use y for yes)"
		
		if($doubleConf -eq "y")
		{
		try{
		Get-WmiObject -Class Win32_volume -Filter "DriveLetter = '$driveLetter'" |Set-WmiInstance -Arguments @{DriveLetter='R:'}
		successText ("CD-ROM Drive changed to R:")
		}
		catch{
		failText("Drive was not changed.")
		}
		}
		else{
		infoText("Exiting drive block")
		}
		}

$newNameConf = read-host "Would you like to rename the PC?(Use y for yes)"

	if($newNameConf -eq 'y')
	{
	$newName = read-host "Please enter the name of the new PC"
		
		if([string]::IsNullOrWhiteSpace($newName) -eq $true)
		{
		failText("Check name format and try again.")
		}
		else{
			infoText("Making sure PC on network doesn't already have the same name..")
			if(test-connection -computerName $newName -quiet)
			{
			errorText("PC with the same name already exists, please pick another name")
			}
				rename-computer -newName $newName
		}
		
	}

successText("Completed. Please remember to license software.")
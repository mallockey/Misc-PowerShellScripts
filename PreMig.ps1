$currentUserProfile = $env:USERPROFILE
$currentUserFolder = split-path $currentUserProfile -leaf

function testPath($path)
{
 $path = test-path $path
	 if($path -eq $true)
	 {
	 return $true
	 }
	 else
	 {
	 return $false
	 }

}

function infoText($output)
{
	$output = write-host -foreGroundColor yellow "INFO:"$output
	return $output
}

function failText($output)
	{
	$output = write-host -ForeGroundColor red "ERROR:"$output
	return $output
	}
function successText($output){
	$output = write-host -foreGroundColor green "SUCCESS:"$output
	return $output
}
function takeScreenShot{

	Param(
	$fileName
	)

	Add-Type -AssemblyName System.Windows.Forms
	Add-type -AssemblyName System.Drawing

	# Gather Screen resolution information
	$Screen = [System.Windows.Forms.SystemInformation]::VirtualScreen

	# Create bitmap using the top-left and bottom-right bounds
	$bitmap = New-Object System.Drawing.Bitmap $Screen.Width, $Screen.Height

	# Create Graphics object
	$graphic = [System.Drawing.Graphics]::FromImage($bitmap)

	# Capture screen
	$graphic.CopyFromScreen($Screen.Left, $Screen.Top, 0, 0, $bitmap.Size)

	# Save to file
	$bitmap.Save($fileName)
	successText( "Screenshot saved to: $fileName")
	

}

function postChecks{

	Param(
	$test,
	$result
	)
	$currentTest = testPath($test)
	if($currentTest -eq $true){
	$result = "$test was backed up successfully"
	}
	else{
	$result = "$test was NOT backed up successfully"
	}
	return $result

}

#START!
$testMailMigrationFolder = testPath("C:\Kits\MailMigration")
if($testMailMigrationFolder -eq $false){
	infoText("Creating C:\Kits\MailMigration Folder")
	write-host("-------------------------")
	new-item -ItemType directory -Path C:\kits\MailMigration
	write-host("-------------------------")
}
else{
write-host "C:\Kits\MailMigration already exists, saving files there"
}

#

#Screenshot Outlook Mail View
start-process outlook.exe 
start-sleep -seconds 10
$fileName = "C:\kits\MailMigration\OutlookView.bmp"
takeScreenShot -fileName $fileName

#Screenshot Outlook Calendar View
start-process outlook.exe -argumentlist "/select outlook:calendar"
start-sleep -seconds 10
$fileName = "C:\kits\MailMigration\OutlookCalendarView.bmp"
takeScreenShot -fileName $fileName

#Screenshot Outlook ContactsView
start-process outlook.exe -argumentlist "/select outlook:contacts"
start-sleep -seconds 10
$fileName = "C:\kits\MailMigration\OutlookContactsView.bmp"
takeScreenShot -fileName $fileName
$currentUserProfile = $env:USERPROFILE
$autoComplete = $currentUserProfile+"\appdata\local\microsoft\outlook\roamcache\"
$signatures = $currentUserProfile+"\appdata\roaming\microsoft\signatures"

$autoCompleteTest = testPath($autoComplete)
if($autoCompleteTest -eq $true){
	successText ("AutoComplete found, backing up to C:\Kits\MailMigration")
	copy-item -Path $autoComplete -destination "C:\Kits\MailMigration" -recurse
}
else{
	infoText("No RoamCache folder found under $autoComplete")

}

$autoCompleteTest = testPath($signatures)
if($autoCompleteTest -eq $true){
	successText ("Signatures found, backing up to C:\Kits\MailMigration")
	copy-item -Path $signatures -destination "C:\Kits\MailMigration" -recurse
}
else{
	infoText("No Signatures folder found under $signatures")

}
write-host "---------------------------------------------"
write-host "Checking for PSTS under $currentUserProfile"
$PSTS = Get-ChildItem "C:\Users\$currentUserFolder" -Recurse -Filter '*.pst' 
$index = 1
	if($PSTS -eq $null){
	write-host "No PSTS found."
	}
	else{
		write-host "------------------------------------"
		write-host "Found some!"
		foreach ($pst in $psts){
		write-host "$index)"$PST.name 
		write-host $PST.directory
		$index++
		}
	}
write-host "----------------------------Results--------------------------------------"
$postCheckAutoComplete = "C:\kits\MailMigration\roamcache\"
$postCheckAutoComplete = postChecks -test $postCheckAutoComplete
write-host "1)"$postCheckAutoComplete


$postCheckSignatures = "C:\kits\MailMigration\Signatures\"
$postCheckSignatures = postChecks -test $postCheckSignatures
write-host "2)"$postCheckSignatures

if($PSTS -ne $null){
write-host "3) PSTS were found please check above folders"
}

write-host "----------------------------End Of Results--------------------------------------"
write-host "Press Enter to exit"
read-host


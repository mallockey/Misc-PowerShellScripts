<#
Author: Josh Melo
Last Updated: 11/28/18
#>
write-output "
-----------------------------------------------------------------------------
This script is used for backing up user specific data during a mail migration.
AutoComplete, Signatures, and screenshots of Outlook Mail tab, 
Outlook Calendar tab, and Outlook Contacts tab are automaticlly taken 
and stored in C:\Kits\MailMigration. Mailbox rules will be displayed in the script and 
exported to a CSV WITHOUT CONDITIONS. PST files will be checked for under the
users folder and the location will be noted.
-----------------------------------------------------------------------------"

start-sleep -seconds 5

$currentUserProfile = $env:USERPROFILE
$mailMigrationFolder = "C:\Kits\MailMigration"

$currentUserFolder = split-path $currentUserProfile -leaf

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


$testMail = testPath($mailMigrationFolder)
	if($testMail -eq $true)
	{
	failText("MailMigration folder already exists in C:\Kits")
	failText("Please rename existing MailMigration folder and rerun script")
	read-host
	exit
	}


function getRules(){
#This code was taken from Scripting Guy blog here:
#https://blogs.technet.microsoft.com/heyscriptingguy/2009/12/15/hey-scripting-guy-how-can-i-tell-which-outlook-rules-i-have-created/

Add-Type -AssemblyName microsoft.office.interop.outlook 
$olFolders = "Microsoft.Office.Interop.Outlook.OlDefaultFolders" -as [type]
$outlook = New-Object -ComObject outlook.application
$namespace  = $Outlook.GetNameSpace("mapi")
$folder = $namespace.getDefaultFolder($olFolders::olFolderInbox)
$rules = $outlook.session.DefaultStore.GetRules()
$rules |
Sort-Object -Property ExecutionOrder |
Format-Table -Property Name, ExecutionOrder, Enabled, isLocalRule -AutoSize
$rules | export-csv "$mailMigrationFolder\Rules.csv"
foreach($rule in $rules){
	if($rule.IsLocalRule -eq $true){
	write-host $rule "is local. Please consult PTM"
	}
}

}
	write-host "---------------------------------------------"
function takeScreenShot{

	Param(
	$fileName
	)
        write-host "Please have Outlook open on screen and maximized."
        write-host "Close any expanded mailboxes/calendars/contacts to allow maximum view"
		write-host "---------------------------------------------"
        $i = 16
        while($i -ge 0)
        {
		 $i--
			if($i -le 15 -and $i -ge 10){
			start-sleep -seconds 1
			write-host -foreGroundColor green "Screenshot in: $i seconds"
			}
				if($i -lt 10 -and $i -ge 5){
				start-sleep -seconds 1
				write-host -foreGroundColor yellow "Screenshot in: $i seconds"
				}
					if($i -lt 5 -and $i -gt 0){
					start-sleep -seconds 1
					write-host -foreGroundColor red "Screenshot in: $i seconds"
					}
			   
					   if($i -eq 0){
					   write-host -foregroundColor red "SCREENSHOT TIME! CLICK!"
					   break
					   }
       
        
        }
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
    try{
	$bitmap.Save($fileName)
	successText( "Screenshot saved to: $fileName")
	write-host "---------------------------------------------"
    }

    catch{
        failText("Taking screenshot, please rerun script.")
        exit
    }
	

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
	#Only in variable to silently make the directory
	$mailDirectory = new-item -ItemType directory -Path C:\kits\MailMigration
	
}
else{
infoText("C:\Kits\MailMigration already exists, saving files there")
}
infoText("Loading Outlook to take screenshots...")
start-sleep -seconds 7
#Screenshot Outlook Mail View
start-process outlook.exe 
$fileName = "$mailMigrationFolder\OutlookView.bmp"
takeScreenShot -fileName $fileName

#Screenshot Outlook Calendar View
start-process outlook.exe -argumentlist "/select outlook:calendar"
$fileName = "$mailMigrationFolder\OutlookCalendarView.bmp"
takeScreenShot -fileName $fileName

#Screenshot Outlook ContactsView
start-process outlook.exe -argumentlist "/select outlook:contacts"
$fileName = "$mailMigrationFolder\OutlookContactsView.bmp"
takeScreenShot -fileName $fileName

$currentUserProfile = $env:USERPROFILE
$autoComplete = $currentUserProfile+"\appdata\local\microsoft\outlook\roamcache\"
$signatures = $currentUserProfile+"\appdata\roaming\microsoft\signatures"

$autoCompleteTest = testPath($autoComplete)
if($autoCompleteTest -eq $true){
	successText ("AutoComplete found, backing up to C:\Kits\MailMigration")
	copy-item -Path $autoComplete -destination "$mailMigrationFolder" -recurse
	write-host "------------------------------------------------------------------"
}
else{
	infoText("No RoamCache folder found under $autoComplete")
}

$autoCompleteTest = testPath($signatures)
if($autoCompleteTest -eq $true){
	successText ("Signatures found, backing up to C:\Kits\MailMigration")
	copy-item -Path $signatures -destination "$mailMigrationFolder" -recurse
	write-host "------------------------------------------------------------------"
}

else{
	infoText("No Signatures folder found under $signatures")
}
write-host "------------------------PST Check------------------------------------"
write-host "Checking for PSTS under $currentUserProfile"
$PSTS = Get-ChildItem "C:\Users\$currentUserFolder" -Recurse -Filter '*.pst' 
$index = 1
	if($PSTS -eq $null){
	write-host "No PSTS found."
	}
	else{
		write-host "---------------------------------------------"
		write-host "Found some!"
            foreach ($pst in $psts){
            write-host "$index)"$PST.name 
            write-host $PST.directory
            $index++
		}
	}

write-host "--------------------------Checking for Rules------------------------------"
start-process outlook.exe
getRules

write-host "--------------------------------Results-----------------------------------------"
$postCheckAutoComplete = "$mailMigrationFolder\roamcache\"
$postCheckAutoComplete = postChecks -test $postCheckAutoComplete
write-host "1)"$postCheckAutoComplete

$postCheckSignatures = "$mailMigrationFolder\Signatures\"
$postCheckSignatures = postChecks -test $postCheckSignatures
write-host "2)"$postCheckSignatures

    if($PSTS -ne $null){
    write-host "3) PSTS were found please check above folders"
    }

write-host "----------------------------End Of Results--------------------------------------"

while($confirmation -ne "*"){
$confirmation = read-host "Enter * to exit"
}

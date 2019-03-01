<#
Author: Josh Melo
Last Updated: 02/28/19
#>
write-output "
-----------------------Mail Migration Prep Script-----------------------
*For use with Office 2010 and above*
This script will do the following:
-Backup AutoComplete
-Record user information in CSVs files
-Backup Signatures
-Backup Rules(Without Conditions)
-Take Screenshots of Outlook tabs
-Check for PSTs under C:\ and note location. Does not backup PSTs
-Record Number of Contacts
------------------------------------------------------------------------"
function testPath{
    Param(
    $path
    )
    $path = test-path $path
        if($path -eq $true){
           return $true
        }
        else{
        return $false
        }
}
function postChecks{
    Param(
    $test,
    $result
    )
    $currentTest = testPath -path $test
	if($currentTest -eq $true){
		$checkIfFolderIsEmpty = Get-ChildItem $test
		if($checkIfFolderIsEmpty -eq $null){
		    $result = "Warning(Folder was empty)"
		}
		else{
		    $result = "Success"
		}
	}
	else{
	$result = "Failed"
	}
return $result
}

function getRules{
#This code was taken from Scripting Guy blog here:
#https://blogs.technet.microsoft.com/heyscriptingguy/2009/12/15/hey-scripting-guy-how-can-i-tell-which-outlook-rules-i-have-created/
$numLocalRules = 0
$numServerRules = 0
$rulesString = ""
    try{
        Add-Type -AssemblyName microsoft.office.interop.outlook 
        $olFolders = "Microsoft.Office.Interop.Outlook.OlDefaultFolders" -as [type]
        $outlook = New-Object -ComObject outlook.application
        $namespace  = $Outlook.GetNameSpace("mapi")
        $folder = $namespace.getDefaultFolder($olFolders::olFolderInbox)
        $rules = $outlook.session.DefaultStore.GetRules()
        $rules | Sort-Object -Property ExecutionOrder |
        Format-Table -Property Name, ExecutionOrder, Enabled, isLocalRule -AutoSize
		foreach($rule in $rules){
		    if($rule.IsLocal -eq $true){
			$numLocalRules++
		    }
		    else{
			$numServerRules++
		    }
		}
	$rulesString = "$numLocalRules Local Rules | $numServerRules Server rules"
        $rules | export-csv "$mailMigrationFolder\Rules.csv" -noTypeInformation
        $postRulesCheck = "$mailMigrationFolder\Rules.csv"
        $postRulesCheck = postChecks -test $postRulesCheck
        $row = $testTable.NewRow()
        $row.BackedUpData = "OutlookRules"
        $row.OriginalPath = "$rulesString"
        $row.BackedUpPath = "$mailMigrationFolder\Rules.csv"
        $row.Pass = $postRulesCheck
        $testTable.Rows.Add($row)
    }
    catch{
        $row = $testTable.NewRow()
        $row.BackedUpData = "OutlookRules"
        $row.OriginalPath = "Not Applicable"
        $row.BackedUpPath = "N/A"
        $row.Pass = "Failed"
        $testTable.Rows.Add($row)   
    }
}
function takeScreenShot{
Param(
$fileName
)
$i = 5
    while($i -gt 0){
      Write-Progress -Activity "Collecting Data" -CurrentOperation "Taking Screenshot in : $i Seconds"
      start-sleep -seconds 1
      $i--
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
    $bitmap.Save($fileName)
    Write-Progress -Activity "Collecting Data" -CurrentOperation "Screenshot Saved to $fileName"
    start-sleep -seconds 2
}
#START!
start-sleep -seconds 4

$testTable = New-Object system.Data.DataTable “$tableName”
$col1 = New-Object system.Data.DataColumn BackedUpData,([string])
$col2 = New-Object system.Data.DataColumn OriginalPath,([string])
$col3 = New-Object system.Data.DataColumn BackedUpPath,([string])
$col4 = New-Object system.Data.DataColumn Pass,([string])
$testTable.columns.add($col1)
$testTable.columns.add($col2)
$testTable.columns.add($col3)
$testTable.columns.add($col4)


$userDomain = $env:userDomain
$userName = $env:username
$PCName = hostname

$row = $testTable.NewRow()
$row.BackedUpData = "Username/Domain/PCName"
$row.OriginalPath = "$userName"
$row.BackedUpPath = "$userDomain"
$row.Pass = "$PCName"
$testTable.Rows.Add($row)
$todaysDate = get-date -Format MM-dd-yyyy
$currentUserProfile = $env:USERPROFILE
$mailMigrationFolder = "C:\Kits\MailMigration_$todaysDate"
$currentUserFolder = split-path $currentUserProfile -leaf

$testMail = testPath -path $mailMigrationFolder
    while($testMail -eq $true){     
       $mailMigrationFolder = "C:\Kits\"
       $mailMigrationFolder += read-host "Enter another name for the migration folder, folder already in use"
       $testMail = testPath -path $mailMigrationFolder
    }

$createMigrationFolder = New-Item -ItemType Directory $mailMigrationFolder
Write-Progress -Activity "Collecting Data" -CurrentOperation "Creating folder: $mailMigrationFolder"
start-sleep -Seconds 2
   
Write-Progress -Activity "Collecting Data" -CurrentOperation "Loading Outlook to take screenshots"
start-sleep -seconds 5
#Screenshot Outlook Mail View
start-process outlook.exe 
read-host "Press Enter when ready to take screenshot"
$fileName = "$mailMigrationFolder\OutlookMailView.bmp"
takeScreenShot -fileName $fileName

#Screenshot Outlook Calendar View
start-process outlook.exe -argumentlist "/select outlook:calendar"
read-host "Press Enter when ready to take screenshot"
$fileName = "$mailMigrationFolder\OutlookCalendarView.bmp"
takeScreenShot -fileName $fileName

#Screenshot Outlook ContactsView
start-process outlook.exe -argumentlist "/select outlook:contacts"
read-host "Press Enter when ready to take screenshot"
$fileName = "$mailMigrationFolder\OutlookContactsView.bmp"
takeScreenShot -fileName $fileName

while($moreScreenShots -ne "n"){
    $moreScreenShots = read-host "Do you want to take more screenshots?(Enter y or n)"
        if($moreScreenShots -eq "n"){
            break
        }
    $fileName = read-host "Enter a file name for this screenshot"
    $fileName = "$mailMigrationFolder\$fileName.bmp"
    takeScreenShot -fileName $fileName
}

Write-Progress -Activity "Collecting Data" -CurrentOperation "Backing up AutoComplete"
Start-Sleep -Seconds 1
$autoComplete = $currentUserProfile+"\appdata\local\microsoft\outlook\roamcache\"
Write-Progress -Activity "Collecting Data" -CurrentOperation "Backing up Signatures"
start-sleep -Seconds 1
$signatures = $currentUserProfile+"\appdata\roaming\microsoft\signatures"

$autoCompleteTest = testPath -path $autoComplete
    if($autoCompleteTest -eq $true){
    copy-item -Path $autoComplete -destination "$mailMigrationFolder" -recurse
    }

$signaturesTest = testPath -path $signatures
    if($signaturesTest -eq $true){
    copy-item -Path $signatures -destination "$mailMigrationFolder" -recurse
    }

Write-Progress -Activity "Collecting Data" -CurrentOperation "Checking for PSTs under C:\"
$PSTS = Get-ChildItem "C:\" -Recurse -Filter '*.pst' -ErrorAction SilentlyContinue
    foreach ($pst in $psts){
        $row = $testTable.NewRow()
        $row.BackedUpData = "PST Name/Location"
        $PSTName = $PST.name
        $PSTDirectory = $PST.directory
        $row.OriginalPath = "$PSTName"
        $row.BackedUpPath = "$PSTDirectory"
        $row.Pass  = "Location Noted Only"
        $testTable.Rows.Add($row) 
    }
Write-Progress -Activity "Collecting Data" -CurrentOperation "Getting Total Contacts"

$row = $testTable.NewRow()
try{
    $numContacts = 0
    $outlook = New-Object -ComObject Outlook.Application
    $contacts = $outlook.session.GetDefaultFolder(10).items
    $contacts | ForEach-Object {$numContacts++}
    $row.BackedUpData = "Contacts"
    $row.OriginalPath = "Total Contacts: $numContacts"
    $row.BackedUpPath = "Not Applicable"
    $row.Pass = "Success"
    $testTable.Rows.Add($row)
}
catch{
    write-progress -Activity "Collecting Data" -currentOperation "Unable to get contacts"
    $row.BackedUpData = "Contacts"
    $row.OriginalPath = "N/A"
    $row.BackedUpPath = "N/A"
    $row.Pass = "Failed"
    $testTable.Rows.Add($row)
}

start-process outlook.exe
Write-Progress -Activity "Collecting Data" -CurrentOperation "Getting Outlook Rules"
getRules
<#
Creates table and validates whether or not data was backed up
#>
$postCheckAutoComplete = "$mailMigrationFolder\roamcache\"
$postCheckAutoComplete = postChecks -test $postCheckAutoComplete
$row = $testTable.NewRow()
$row.BackedUpData = "AutoComplete"
$row.OriginalPath = "$autoComplete"
$row.BackedUpPath = "$mailMigrationFolder\roamcache\"
$row.Pass = $postCheckAutoComplete
$testTable.Rows.Add($row)

$postCheckSignatures = "$mailMigrationFolder\Signatures\"
$postCheckSignatures = postChecks -test $postCheckSignatures
$row = $testTable.NewRow()
$row.BackedUpData = "Signatures"
$row.OriginalPath = "$signatures"
$row.BackedUpPath = "$mailMigrationFolder\signatures\"
$row.Pass = $postCheckSignatures
$testTable.Rows.Add($row)

$testTable | format-table -AutoSize 
$testTable | export-csv "$mailMigrationFolder\TestsTable.csv" -NoTypeInformation

write-host "All data was backed up to: $mailMigrationFolder"
write-host "------------------------------------------------------------------------"
read-host "Press Enter to exit"

<#
Author: Josh Melo
Last Updated: 02/17/19
#>
write-output "
-----------------------Mail Migration Prep Script-----------------------
This script will do the following:
-Backup AutoComplete
-Record user information in CSVs files
-Backup Signatures
-Backup Rules(Without Conditions)
-Take Screenshots of Outlook tabs
-Check for PSTs under C:\
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
    try{
        Add-Type -AssemblyName microsoft.office.interop.outlook 
        $olFolders = "Microsoft.Office.Interop.Outlook.OlDefaultFolders" -as [type]
        $outlook = New-Object -ComObject outlook.application
        $namespace  = $Outlook.GetNameSpace("mapi")
        $folder = $namespace.getDefaultFolder($olFolders::olFolderInbox)
        $rules = $outlook.session.DefaultStore.GetRules()
        $rules | Sort-Object -Property ExecutionOrder |
        Format-Table -Property Name, ExecutionOrder, Enabled, isLocalRule -AutoSize

        $rules | export-csv "$mailMigrationFolder\Rules.csv" -noTypeInformation
        $postRulesCheck = test-path -Type leaf "$mailMigrationFolder\Rules.csv"
        $row = $testTable.NewRow()
        $row.BackedUpData = "OutlookRules"
        $row.OriginalPath = "Not Applicable"
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

$infoTable = New-Object system.Data.DataTable “$infoTable”
$inCol1 = New-Object system.Data.DataColumn Info,([string])
$inCol2 = New-Object system.Data.DataColumn Value1,([string])
$inCol3 = New-Object system.Data.DataColumn Value2,([string])
$inCol4 = New-Object system.Data.DataColumn Value3,([string])
$infoTable.columns.add($inCol1)
$infoTable.columns.add($inCol2)
$infoTable.columns.add($inCol3)
$infoTable.columns.add($inCol4)

$userDomain = $env:userDomain
$userName = $env:username
$PCName = hostname

$row = $infoTable.NewRow()
$row.Info = "UserInfo"
$row.Value1 = "$userName"
$row.Value2 = "$userDomain"
$row.Value3 = "$PCName"
$infoTable.Rows.Add($row)
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
        $row = $infoTable.NewRow()
        $row.Info = "PST"
        $PSTName = $PST.name
        $PSTDirectory = $PST.directory
        $row.Value1 = "$PSTName"
        $row.Value2 = "$PSTDirectory"
        $infoTable.Rows.Add($row) 
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

$userString = $userName + "On_" + $PCName

$infoTable | format-table 
$infoTable |  export-csv "$mailMigrationFolder\$userString.csv" -NoTypeInformation

$testTable | format-table 
$testTable | export-csv "$mailMigrationFolder\TestsTable.csv" -NoTypeInformation

write-host "All data was backed up to: $mailMigrationFolder"
write-host "------------------------------------------------------------------------"
read-host "Press Enter to exit"

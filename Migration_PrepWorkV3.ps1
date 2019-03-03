<#
Author: Josh Melo
Last Updated: 03/2/19
-Changed:
Data is now exported to list instead of table, easier to read and more accurate.
#>
Write-Output "
-----------------------Mail Migration Prep Script-----------------------
*For use with Office 2010 and above*
This script will do the following:
-Backup AutoComplete
-Backup Signatures
-Backup Rules(Without Conditions)
-Take Screenshots of Outlook tabs
-Check for PSTs under C:\ and note location. Does not backup PSTs
-Record Number of Contacts
-Exports List at end to Results.txt
------------------------------------------------------------------------"
function createList {
  param(
    $arrayKey,
    $arrayValue
  )
  $tempObj = New-Object -TypeName PSObject

  for ($i = 0; $i -lt $arrayKey.count; $i++) {
    $tempObj | Add-Member -MemberType NoteProperty -Name $arrayKey[$i] -Value $arrayValue[$i]
  }
  $arrayOfInfo.Add($tempObj) | Out-Null
}
function testPath {
  param(
    $path
  )
  $path = Test-Path $path
  if ($path -eq $true) {
    return $true
  }
  else {
    return $false
  }
}
function postChecks {
  param(
    $test,
    $result
  )
  $currentTest = testPath -Path $test
  if ($currentTest -eq $true) {
    $checkIfFolderIsEmpty = Get-ChildItem $test
    if ($checkIfFolderIsEmpty -eq $null) {
      $result = "Warning(Folder was empty)"
    }
    else {
      $result = "Success"
    }
  }
  else {
    $result = "Failed"
  }
  return $result
}
function getRules {
  #This code was taken from Scripting Guy blog here:
  #https://blogs.technet.microsoft.com/heyscriptingguy/2009/12/15/hey-scripting-guy-how-can-i-tell-which-outlook-rules-i-have-created/
  $numLocalRules = 0
  $numServerRules = 0
  $rulesString = ""
  try {
    Add-Type -AssemblyName microsoft.office.interop.outlook
    $olFolders = "Microsoft.Office.Interop.Outlook.OlDefaultFolders" -as [type]
    $outlook = New-Object -ComObject outlook.application
    $namespace = $Outlook.GetNameSpace("mapi")
    $folder = $namespace.getDefaultFolder($olFolders::olFolderInbox)
    $rules = $outlook.session.DefaultStore.GetRules()
    $rules | Sort-Object -Property ExecutionOrder |
    Format-Table -Property Name,ExecutionOrder,Enabled,isLocalRule -AutoSize
    foreach ($rule in $rules) {
      if ($rule.IsLocal -eq $true) {
        $numLocalRules++
      }
      else {
        $numServerRules++
      }
    }
    $rulesString = "$numLocalRules Local Rules | $numServerRules Server rules"
    $rules | Export-Csv "$mailMigrationFolder\Rules.csv" -NoTypeInformation
    $postRulesCheck = "$mailMigrationFolder\Rules.csv"
    $postRulesCheck = postChecks -test $postRulesCheck
    $rulesInfo = "Test","Local Rules","Server Rules","Status"
    $rulesValues = "Rules","$numLocalRules","$numServerRules","$postRulesCheck"
  }
  catch {
    $rulesInfo = "Test","Local Rules","Server Rules","Status"
    $rulesValues = "Rules","N/A","N/A","Failed"
    createList -arrayKey $rulesInfo -arrayValue $rulesValues
  }
}
function takeScreenShot {
  param(
    $fileName
  )
  $i = 5
  while ($i -gt 0) {
    Write-Progress -Activity "Collecting Data" -CurrentOperation "Taking Screenshot in : $i Seconds"
    Start-Sleep -Seconds 1
    $i --
  }
  Add-Type -AssemblyName System.Windows.Forms
  Add-Type -AssemblyName System.Drawing
  # Gather Screen resolution information
  $Screen = [System.Windows.Forms.SystemInformation]::VirtualScreen
  # Create bitmap using the top-left and bottom-right bounds
  $bitmap = New-Object System.Drawing.Bitmap $Screen.Width,$Screen.Height
  # Create Graphics object
  $graphic = [System.Drawing.Graphics]::FromImage($bitmap)
  # Capture screen
  $graphic.CopyFromScreen($Screen.Left,$Screen.Top,0,0,$bitmap.Size)
  # Save to file
  $bitmap.Save($fileName)
  Write-Progress -Activity "Collecting Data" -CurrentOperation "Screenshot Saved to $fileName"
  Start-Sleep -Seconds 2
}
#START!
Start-Sleep -Seconds 4
$arrayOfInfo = [System.Collections.ArrayList]@()
$PCName = hostname
$userInfo = "$env:username","$env:userDomain"
$userInfo += $PCName
$userHeaders = "Username","Domain","PCName"
createList -arrayKey $userHeaders -arrayValue $userInfo

$todaysDate = Get-Date -Format MM-dd-yyyy
$currentUserProfile = $env:USERPROFILE
$mailMigrationFolder = "C:\Kits\MailMigration_$todaysDate"

$testMail = testPath -Path $mailMigrationFolder
while ($testMail -eq $true) {
  $mailMigrationFolder = "C:\Kits\"
  $mailMigrationFolder += Read-Host "Enter another name for the migration folder, folder already in use"
  $testMail = testPath -Path $mailMigrationFolder
}

$createMigrationFolder = New-Item -ItemType Directory $mailMigrationFolder
Write-Progress -Activity "Collecting Data" -CurrentOperation "Creating folder: $mailMigrationFolder"
Start-Sleep -Seconds 2

Write-Progress -Activity "Collecting Data" -CurrentOperation "Loading Outlook to take screenshots"
Start-Sleep -Seconds 5
#Screenshot Outlook Mail View
Start-Process outlook.exe
Read-Host "Press Enter when ready to take screenshot"
$fileName = "$mailMigrationFolder\OutlookMailView.bmp"
takeScreenShot -FileName $fileName

#Screenshot Outlook Calendar View
Start-Process outlook.exe -ArgumentList "/select outlook:calendar"
Read-Host "Press Enter when ready to take screenshot"
$fileName = "$mailMigrationFolder\OutlookCalendarView.bmp"
takeScreenShot -FileName $fileName

#Screenshot Outlook ContactsView
Start-Process outlook.exe -ArgumentList "/select outlook:contacts"
Read-Host "Press Enter when ready to take screenshot"
$fileName = "$mailMigrationFolder\OutlookContactsView.bmp"
takeScreenShot -FileName $fileName

while ($moreScreenShots -ne "n") {
  $moreScreenShots = Read-Host "Do you want to take more screenshots?(Enter y or n)"
  if ($moreScreenShots -eq "n") {
    break
  }
  $fileName = Read-Host "Enter a file name for this screenshot"
  $fileName = "$mailMigrationFolder\$fileName.bmp"
  takeScreenShot -FileName $fileName
}

Write-Progress -Activity "Collecting Data" -CurrentOperation "Backing up AutoComplete"
Start-Sleep -Seconds 1
$autoComplete = $currentUserProfile + "\appdata\local\microsoft\outlook\roamcache\"
Write-Progress -Activity "Collecting Data" -CurrentOperation "Backing up Signatures"
Start-Sleep -Seconds 1
$signatures = $currentUserProfile + "\appdata\roaming\microsoft\signatures"

$autoCompleteTest = testPath -Path $autoComplete
if ($autoCompleteTest -eq $true) {
  Copy-Item -Path $autoComplete -Destination "$mailMigrationFolder" -Recurse
}
$signaturesTest = testPath -Path $signatures
if ($signaturesTest -eq $true) {
  Copy-Item -Path $signatures -Destination "$mailMigrationFolder" -Recurse
}
Write-Progress -Activity "Collecting Data" -CurrentOperation "Checking for PSTs under C:\"
$PSTS = Get-ChildItem "C:\" -Recurse -Filter '*.pst' -ErrorAction SilentlyContinue
foreach ($pst in $psts) {
  $PSTName = $pst.Name
  $PSTDirectory = $pst.Directory
  $PSTKeyInfo = "PSTName","Path"
  $PSTValueInfo = "$PSTName","$PSTDirectory"
  createList -arrayKey $PSTKeyInfo -arrayValue $PSTValueInfo
}
Write-Progress -Activity "Collecting Data" -CurrentOperation "Getting Total Contacts"
try {
  $numContacts = 0
  $outlook = New-Object -ComObject Outlook.Application
  $contacts = $outlook.session.getDefaultFolder(10).items
  $contacts | ForEach-Object { $numContacts++ }
  $contactsKeyInfo = "Contacts","Status"
  $contactsValueInfo = "$numContacts","Success"
  createList -arrayKey $contactsKeyInfo -arrayValue $contactsValueInfo
}
catch {
  Write-Progress -Activity "Collecting Data" -CurrentOperation "Unable to get contacts"
  $contactsKeyInfo = "Contacts","Status"
  $contactsValueInfo = "Not Available","Failed"
  createList -arrayKey $contactsKeyInfo -arrayValue $contactsValueInfo
}

Start-Process outlook.exe
Write-Progress -Activity "Collecting Data" -CurrentOperation "Getting Outlook Rules"
getRules

$postCheckAutoComplete = "$mailMigrationFolder\roamcache\"
$postCheckAutoComplete = postChecks -test $postCheckAutoComplete
$autoCompleteInfo = "Test","Status"
$autoCompleteValues = "AutoComplete","$postCheckAutoComplete"

createList -arrayKey $autoCompleteInfo -arrayValue $autoCompleteValues

$postCheckSignatures = "$mailMigrationFolder\Signatures\"
$postCheckSignatures = postChecks -test $postCheckSignatures
$signatureInfo = "Test","Status"
$signatureValues = "Signatures","$postCheckSignatures"

createList -arrayKey $signatureInfo -arrayValue $signatureValues
foreach ($obj in $arrayOfInfo) {
  ($obj | Format-List | Out-String).Trim() + "`n" | Out-File "$mailMigrationFolder\Results.txt" -Append -Encoding unicode
}

$arrayOfInfo | Format-List

write-host "All data was backed up to: $mailMigrationFolder"
write-host "------------------------------------------------------------------------"
read-host "Press Enter to exit"

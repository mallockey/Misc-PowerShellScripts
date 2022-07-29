param(
  [Switch]$LocalRulesOnly
)
<#
The following data does not exist in within the COM object therefore we cannot query its data:
  Condition 21 - ("Message size is between x and y...")
  Condition 7 - (Message is marked with the specified level of sensitivity.)
  Condition 8 - (Message is flagged for specific response)
#>

$RulesActions = @{
  #https://docs.microsoft.com/en-us/office/vba/api/outlook.olruleactiontype
  1  = 'Move to the specified folder.'
  2  = 'Assign categories to the message.'
  3  = 'Delete the message.'
  4  = 'Permanently delete the message.'
  5  = 'Copy the message to a specified folder'
  6  = 'Forward the message to the specified recipients'
  7  = 'Forward the message as an attachment to the specified recipients'
  8  = 'Redirect the message to the specified recipients'
  9  = 'Request the server to reply with the specified mail item.'
  10 = 'Use the specified template (.oft) file as a form template.'
  11 = 'Flag the message for action in the specified number of days:'
  12 = 'Flag the message with a specified colored flag.'
  13 = 'Clear the message flag.'
  14 = 'Mark the message with the specified level of importance.'
  15 = 'Mark the message with the specified level of sensitivity.'
  16 = 'Print the message on the default printer.'
  17 = 'Play a sound file.'
  18 = 'Run an .exe file.'
  19 = 'Request the server to reply with the specified mail item.'
  20 = 'Run a script.'
  21 = 'Stop processing more rules.'
  22 = 'Perform a custom action.'
  23 = 'Display the specified text in the New Item Alert dialog box.'
  24 = 'Display a desktop alert.'
  25 = 'Request read notification for the message being sent.'
  26 = 'Request delivery notification for the message being sent.'
  27 = 'CC the message to specified recipients.'
  28 = 'Defer delivery of the message by the specified number of minutes.'
  29 = ''
  30 = 'Clear all the categories assigned to the message.'
  31 = ''
  32 = ''
  33 = ''
  34 = ''
  35 = ''
  36 = ''
  37 = ''
  38 = ''
  39 = ''
  40 = ''
  41 = 'Mark the message as a task.'
}

$RuleConditions = @{
  #https://docs.microsoft.com/en-us/office/vba/api/outlook.olruleconditiontype
  1  = 'If Sender is: '
  2  = 'If Subject contains: '
  3  = 'Account is the account specified in AccountRuleCondition.Account.'
  4  = 'Message is sent only to me.'
  5  = 'My name is in the To box.'
  6  = 'Message is marked with the specified level of importance.'
  7  = 'Message is marked with the specified level of sensitivity.'
  8  = 'Message is flagged for the specified action.'
  9  = 'Message has my name in the Cc box.'
  10 = 'Message has my name in the To or Cc box.'
  11 = 'Message does not have my name in the To box.'
  12 = "Sent $('from'), and $('to') where from and to fields are specified:"
  13 = 'Body contains words specified:'
  14 = 'Body or subject contains words specified by:'
  15 = 'If Message header contains: '
  16 = 'Recipient address contains words specified:'
  17 = 'Sender address contains words specified:'
  18 = 'Category is the category specified in:'
  19 = 'Message is an out-of-office message.'
  20 = 'Message has one or more attachments.'
  21 = 'Message size is between x and y in units of KB, where x and y are Integer values.'
  22 = 'Message was received between x and y, where x and y are Date values.'
  23 = 'Message uses the form specified:'
  24 = 'Document property is exactly, contains, or does not contain specified properties.'
  25 = 'Sender is in the address list specified in AddressRuleCondition.Address.'
  26 = 'Message is a meeting invitation or update.'
  27 = 'Rule can run only on the local machine.'
  28 = 'Rule can run only on a specific machine that is not the current machine.'
  29 = 'Message is assigned to any category.'
  30 = 'Message is generated from a specific RSS subscription.'
  31 = 'Message is generated from any RSS subscription.'
}

$RuleType = @{
  #https://docs.microsoft.com/en-us/office/vba/api/outlook.olruletype
  0 = 'When Receiving message'
  1 = 'When Sending message'
}

$ArrayOfRules = [System.Collections.ArrayList]@()

try {
  Add-Type -AssemblyName microsoft.office.interop.outlook 
  $OlFolders = 'Microsoft.Office.Interop.Outlook.OlDefaultFolders' -as [type]
  $Outlook = New-Object -ComObject outlook.application
  $Namespace = $Outlook.GetNameSpace('mapi')
  $Folder = $Namespace.getDefaultFolder($olFolders::olFolderInbox)
  $Rules = $Outlook.session.DefaultStore.GetRules()
}
catch {
  Write-Warning 'There was an error querying Outlook'
}

foreach ($rule in $rules) {
  $AccountSpecified = ''
  $Address = ''
  $AddressRule = ''
  $Categories = ''
  $ConvertedAddress = ''
  $FormName = ''
  $ImportanceLevel = ''
  $Newaddress = ''
  $NonConvertedAddress = ''
  $TempActions = ''
  $TempConditions = ''
  $TempExceptions = ''
  $Text = ''

  if ($LocalRulesOnly) {
    if ($Rule.IsLocalRule -eq $True) {
      #Get Rule Info
    }
    else {
      continue
    }
  }

  $ActionType = $Rule | ForEach-Object { $_.Actions
  } | Where-Object { $_.Enabled -eq $True
  } | Select-Object -ExpandProperty ActionType

  $RuleSteps = 1
  foreach ($Action in $ActionType) {
    $TempActions += "Step $RuleSteps) $($RulesActions[$Action])"
    $RuleSteps++
  }

  $ImportanceLevel += $Rule | ForEach-Object { $_.conditions
  } | Where-Object { $_.Enabled -eq $True
  } | ForEach-Object { $_.Importance }

  $Folder = $Rule | ForEach-Object { $_.Actions
  } | Where-Object { $_.Enabled -eq $True
  } | ForEach-Object { $_.Folder
  } | Select-Object -ExpandProperty FullFolderPath

  $Address = $Rule | ForEach-Object { $_.Conditions
  } | Where-Object { $_.Enabled -eq $True
  } | ForEach-Object { $_.Recipients
  } | ForEach-Object { $_.Address }

  foreach ($Add in $Address) {
    if ($Add -like '*/o=*') {
      $ConvertedAddress += $Rule | ForEach-Object { $_.Conditions
      } | Where-Object { $_.Enabled -eq $True
      } | ForEach-Object { $_.Recipients } | Where-Object { $_.Address -like $Add
      } | Select-Object -ExpandProperty Name
    }
    else {
      $NonConvertedAddress += $Add
    }
  }

  [String]$Newaddress = $ConvertedAddress + $NonConvertedAddress

  $ConditionType = $Rule | ForEach-Object { $_.Conditions
  } | Where-Object { $_.Enabled -eq $True
  } | ForEach-Object { $_.ConditionType }

  $Text += $Rule | ForEach-Object { $_.Conditions
  } | Where-Object { $_.Enabled -eq $True
  } | ForEach-Object { $_.Text }

  $Categories += $Rule | ForEach-Object { $_.Actions
  } | Where-Object { $_.Enabled -eq $True
  } | ForEach-Object { $_.Categories }

  $AccountSpecified += $Rule | ForEach-Object { $_.Conditions
  } | Where-Object { $_.ConditionType -eq 3
  } | ForEach-Object { $_.Account
  } | ForEach-Object { $_.SMTPAddress }

  $FormName += $Rule | ForEach-Object { $_.Conditions
  } | Where-Object { $_.ConditionType -eq 23
  } | ForEach-Object { $_.FormName }

  $AddressRule += $Rule | ForEach-Object { $_.Conditions } | Where-Object { $_.enabled -eq $True } | Where-Object { $_.ConditionType -eq 25 } | ForEach-Object { $_.AddressList.Name }

  foreach ($Condition in $ConditionType) {
    $TempConditions += $RuleConditions[$Condition]
  }
  $ExceptionsID = $Rule | ForEach-Object { $_.Exceptions
  } | Where-Object { $_.Enabled -eq $True
  } | ForEach-Object { $_.ConditionType
  }

  foreach ($exception in $exceptionsID) {
    $exceptionsText = $rule | ForEach-Object { $_.Exceptions
    } | Where-Object { $_.Enabled -eq $true
    } | ForEach-Object { $_.text }
    $TempExceptions += "$($RuleConditions[$exception])  $exceptionsText"
  }

  $TempRuleObj = New-Object -TypeName PSObject 
  $TempRuleObj | Add-Member -MemberType NoteProperty -Name RuleName -Value $Rule.Name
  $TempRuleObj | Add-Member -MemberType NoteProperty -Name RuleType -Value $RuleType[$Rule.RuleType]
  $TempRuleObj | Add-Member -MemberType NoteProperty -Name Conditions -Value $tempConditions
  $TempRuleObj | Add-Member -MemberType NoteProperty -Name RecipientList -Value $newaddress
  $TempRuleObj | Add-Member -MemberType NoteProperty -Name SubjectOrMessage -Value $text
  $TempRuleObj | Add-Member -MemberType NoteProperty -Name Action -Value $tempActions
  $TempRuleObj | Add-Member -MemberType NoteProperty -Name SpecifiedFolder -Value $folder

  if ($FormName -ne '') {
    $TempRuleObj | Add-Member -MemberType NoteProperty -Name Formname -Value $FormName
  }
  if ($AccountSpecified -ne '') {
    $TempRuleObj | Add-Member -MemberType NoteProperty -Name AccountSpecified -Value $AccountSpecified 
  }
  if ($AddressRule -ne '') {
    $TempRuleObj | Add-Member -MemberType NoteProperty -Name DistributionGroup -Value $AddressRule
  }
  if ($Categories -ne '') {
    $TempRuleObj | Add-Member -MemberType NoteProperty -Name Category -Value $Categories
  }
  if ($importanceLevel -ne '') {
    $TempRuleObj | Add-Member -MemberType NoteProperty -Name ImportanceLevel -Value $importanceLevel
  }
  $TempRuleObj | Add-Member -MemberType NoteProperty -Name Exceptions -Value $tempExceptions
  if ($LocalRulesOnly -eq $false) {
    $tempRuleObj | Add-Member -MemberType NoteProperty -Name IsLocalRule -Value "$($rule.IsLocalRule)"
  }
  $TempRuleObj | Add-Member -MemberType NoteProperty -Name Enabled? -Value "$($rule.Enabled)"
  $ArrayOfRules.Add($TempRuleObj) | Out-Null
}


$ArrayOfRules | Format-Table
$ArrayOfRules | Export-Csv 'C:\Kits\Rules.csv' -NoTypeInformation

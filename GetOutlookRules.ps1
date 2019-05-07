Param(
  [switch]$LocalRulesOnly
)
<#
Author: Josh Melo
Last Updated: 5/6/19
-Actions/Conditions have been slightly modified for better readabillity.The original data sets are located in the links 
for each hash table.
The following data does not exist in within the COM object therefore we cannot query its data:
  Condition 21 - ("Message size is between x and y...")
  Condition 7 - (Message is marked with the specified level of sensitivity.)
  Condition 8 - (Message is flagged for specific response)
#>
$rulesActions = @{
  #https://docs.microsoft.com/en-us/office/vba/api/outlook.olruleactiontype
  1 = "Move to the specified folder."
  2 = "Assign categories to the message."
  3 = "Delete the message."
  4 = "Permanently delete the message."
  5 = "Copy the message to a specified folder"
  6 = "Forward the message to the specified recipients"
  7 = "Forward the message as an attachment to the specified recipients"
  8 = "Redirect the message to the specified recipients"
  9 = "Request the server to reply with the specified mail item."
  10 = "Use the specified template (.oft) file as a form template."
  11 = "Flag the message for action in the specified number of days:"
  12 = "Flag the message with a specified colored flag."
  13 = "Clear the message flag."
  14 = "Mark the message with the specified level of importance."
  15 = "Mark the message with the specified level of sensitivity."
  16 = "Print the message on the default printer."
  17 = "Play a sound file."
  18 = "Run an .exe file."
  19 = "Request the server to reply with the specified mail item."
  20 = "Run a script."
  21 = "Stop processing more rules."
  22 = "Perform a custom action."
  23 = "Display the specified text in the New Item Alert dialog box."
  24 = "Display a desktop alert."
  25 = "Request read notification for the message being sent."
  26 = "Request delivery notification for the message being sent."
  27 = "CC the message to specified recipients."
  28 = "Defer delivery of the message by the specified number of minutes."
  29 = ""
  30 = "Clear all the categories assigned to the message."
  31 = ""
  32 = ""
  33 = ""
  34 = ""
  35 = ""
  36 = ""
  37 = ""
  38 = ""
  39 = ""
  40 = ""
  41 = "Mark the message as a task."
}
$ruleConditions = @{
  #https://docs.microsoft.com/en-us/office/vba/api/outlook.olruleconditiontype
  1 = "If Sender is: "
  2 = "If Subject contains: "
  3 = "Account is the account specified in AccountRuleCondition.Account."
  4 = "Message is sent only to me."
  5 = "My name is in the To box."
  6 = "Message is marked with the specified level of importance."
  7 = "Message is marked with the specified level of sensitivity."
  8 = "Message is flagged for the specified action."
  9 = "Message has my name in the Cc box."
  10 = "Message has my name in the To or Cc box."
  11 = "Message does not have my name in the To box."
  12 = "Sent $("from"), and $("to") where from and to fields are specified:"
  13 = "Body contains words specified:"
  14 = "Body or subject contains words specified by:"
  15 = "If Message header contains: "
  16 = "Recipient address contains words specified:"
  17 = "Sender address contains words specified:"
  18 = "Category is the category specified in:"
  19 = "Message is an out-of-office message."
  20 = "Message has one or more attachments."
  21 = "Message size is between x and y in units of KB, where x and y are Integer values."
  22 = "Message was received between x and y, where x and y are Date values."
  23 = "Message uses the form specified:"
  24 = "Document property is exactly, contains, or does not contain specified properties."
  25 = "Sender is in the address list specified in AddressRuleCondition.Address."
  26 = "Message is a meeting invitation or update."
  27 = "Rule can run only on the local machine."
  28 = "Rule can run only on a specific machine that is not the current machine."
  29 = "Message is assigned to any category."
  30 = "Message is generated from a specific RSS subscription."
  31 = "Message is generated from any RSS subscription."
}
$ruleType = @{
  #https://docs.microsoft.com/en-us/office/vba/api/outlook.olruletype
  0 = "When Receiving message"
  1 = "When Sending message"
}
$arrayOfRules = [System.Collections.ArrayList]@()
function getRules{
  Param(
    $LocalRulesOnly
  )

  Add-Type -AssemblyName microsoft.office.interop.outlook 
  $olFolders = "Microsoft.Office.Interop.Outlook.OlDefaultFolders" -as [type]
  $outlook = New-Object -ComObject outlook.application
  $namespace = $Outlook.GetNameSpace("mapi")
  $folder = $namespace.getDefaultFolder($olFolders::olFolderInbox)
  $rules = $outlook.session.DefaultStore.GetRules()

  foreach($rule in $rules){
    $accountSpecified = ""
    $address = ""
    $addressRule = ""
    $categories = ""
    $convertedAddress =""
    $formName = ""
    $importanceLevel = ""
    $newaddress = ""
    $nonConvertedAddress =""
    $tempActions = ""
    $tempConditions = ""
    $tempExceptions = ""
    $text = ""
    
    $RuleName = $rule.Name
    $tempRuleType = $ruleType[$rule.RuleType]
    
    if($LocalRulesOnly){
      if($rule.IsLocalRule -eq $true){
        #Get Rule Info
      }
      else{
        continue
      }
    }
    $actionType = $rule | ForEach-Object{$_.Actions
    } | Where-Object {$_.Enabled -eq $true
    } | Select-Object -ExpandProperty ActionType

    $ruleSteps = 1
    foreach($action in $actionType){
      $tempActions += "Step $ruleSteps) $($rulesActions[$action])"
      $ruleSteps++
    }

    $importanceLevel += $rule | ForEach-Object {$_.conditions
    } | Where-Object {$_.Enabled -eq $true
    } | ForEach-Object {$_.Importance}

    $folder = $rule | ForEach-Object{$_.Actions
    } | Where-Object {$_.Enabled -eq $true
    } | ForEach-Object {$_.Folder
    } | Select-Object -ExpandProperty FullFolderPath
    
    $address = $rule | ForEach-Object {$_.Conditions
    } | Where-Object {$_.Enabled -eq $true
    } | ForEach-Object {$_.Recipients
    } | ForEach-Object {$_.Address}
  
    foreach($add in $address){
      if($add -like "*/o=*"){
        $convertedAddress += $rule | ForEach-Object {$_.Conditions
        } | Where-Object {$_.Enabled -eq $true
        } | ForEach-Object {$_.Recipients} | Where-Object {$_.Address -like $add
        } | Select-Object -ExpandProperty Name
      }
      else{
        $nonConvertedAddress += $add
      }
    }
    [String]$newaddress = $convertedAddress + $nonConvertedAddress

    $conditionType = $rule | ForEach-Object {$_.Conditions
    } | Where-Object {$_.Enabled -eq $true
    } | ForEach-Object {$_.ConditionType}

    $text += $rule | ForEach-Object {$_.Conditions
    } | Where-Object {$_.Enabled -eq $true
    } | ForEach-Object {$_.Text}

    $categories += $rule | ForEach-Object {$_.Actions
    } | Where-Object {$_.Enabled -eq $true
    } | ForEach-Object{$_.Categories}

    $accountSpecified += $rule | ForEach-Object {$_.Conditions
    } | Where-Object {$_.ConditionType -eq 3
    } | ForEach-Object {$_.Account
    } | ForEach-Object {$_.SMTPAddress}

    $formName += $rule | ForEach-Object {$_.Conditions
    } | Where-Object {$_.ConditionType -eq 23
    } | ForEach-Object {$_.FormName}
  
    $addressRule += $rule | ForEach-Object {$_.Conditions
    } | Where-Object {$_.enabled -eq $true
    } | Where-Object {$_.ConditionType -eq 25
    } | ForEach-Object {$_.AddressList.Name}
      
    foreach($condition in $conditionType){
      $tempConditions += $ruleConditions[$condition]
    }
  $exceptionsID = $rule | ForEach-Object {$_.Exceptions
    } | Where-Object {$_.Enabled -eq $true
    } | ForEach-Object {$_.ConditionType
    }
    foreach($exception in $exceptionsID){
      $exceptionsText = $rule | ForEach-Object {$_.Exceptions
      } | Where-Object {$_.Enabled -eq $true
      } | ForEach-Object {$_.text}
      $tempExceptions += "$($ruleConditions[$exception])  $exceptionsText"
    }
    $tempRuleObj = New-Object -TypeName PSObject 
    $tempRuleObj | Add-Member -MemberType NoteProperty -Name RuleName -Value $RuleName
    $tempRuleObj | Add-Member -MemberType NoteProperty -Name RuleType -Value $tempRuleType
    $tempRuleObj | Add-Member -MemberType NoteProperty -Name Conditions -Value $tempConditions
    $tempRuleObj | Add-Member -MemberType NoteProperty -Name RecipientList -Value $newaddress
    $tempRuleObj | Add-Member -MemberType NoteProperty -Name SubjectOrMessage -Value $text
    $tempRuleObj | Add-Member -MemberType NoteProperty -Name Action -Value $tempActions
    $tempRuleObj | Add-Member -MemberType NoteProperty -Name SpecifiedFolder -Value $folder
    if($formName -ne ""){
      $tempRuleObj | Add-Member -MemberType NoteProperty -Name Formname -Value $formName
    }
    if($accountSpecified -ne ""){
      $tempRuleObj | Add-Member -MemberType NoteProperty -Name AccountSpecified  -Value $accountSpecified 
    }
    if($addressRule -ne ""){
      $tempRuleObj | Add-Member -MemberType NoteProperty -Name DistributionGroup -Value $addressRule
    }
    if($categories -ne ""){
      $tempRuleObj | Add-Member -MemberType NoteProperty -Name Category -Value $categories
    }
    if($importanceLevel -ne ""){
      $tempRuleObj | Add-Member -MemberType NoteProperty -Name ImportanceLevel -Value $importanceLevel
    }
    $tempRuleObj | Add-Member -MemberType NoteProperty -Name Exceptions -Value $tempExceptions
    if($LocalRulesOnly -eq $false){
      $tempRuleObj | Add-Member -MemberType NoteProperty -Name IsLocalRule -Value "$($rule.IsLocalRule)"
    }
    $tempRuleObj | Add-Member -MemberType NoteProperty -Name Enabled? -Value "$($rule.Enabled)"
    $arrayOfRules.Add($tempRuleObj) | out-null
    }
}
getRules -LocalRulesOnly $LocalRulesOnly
$arrayOfRules | Format-Table
$arrayOfRules | Export-Csv "C:\Kits\Rules.csv" -NoTypeInformation

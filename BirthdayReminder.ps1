$birthdayCSV = Import-Csv "PathtoCSV"
$todaysDate = Get-Date
$fromNumber = "EnterTwilioNumberHere"
$toNumber = "Your number" 
$accountSID = 'From Twilio'
$authToken = 'From Twilio'
$arrayOfBirthdays = [System.Collections.ArrayList]@()
[Boolean]$upcomingBirthdays = $false

foreach($birthday in $birthdayCSV){
  $currentName = $birthday.Name
  $currentName = $currentName+"'s"
  $currentBirthday = (Get-Date $birthday.Day -Format MM/dd)
  $differenceInDays = New-TimeSpan -Start $todaysDate -End $currentBirthday | select -expandproperty Days
  if($differenceInDays -lt 0 ){
    continue
  }
  if($differenceInDays -lt 30){
    $upcomingBirthdays = $true
    $birthdayString = "$currentName birthday is on $currentBirthday ! `n"
    $arrayOfBirthdays.Add($birthdayString) | Out-Null
  }
}
if($upcomingBirthdays -eq $false){
  $birthdayString = "There are no birthdays in the next 30 days"
  $arrayOfBirthdays.Add($birthdayString) | Out-Null
}
  C:\Scripts\BirthdayReminder\Send-TwilioSMS.ps1 -AccountSID $accountSID -AuthToken $authToken -FromNumber $fromNumber -ToNumber $toNumber -Message $arrayOfBirthdays
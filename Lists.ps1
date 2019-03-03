$arrayOfInfo = [System.Collections.ArrayList]@()
function createList{
Param(
  $arrayKey,
  $arrayValue
)  
$tempObj = New-Object -TypeName PSObject 

  for($i=0; $i -lt $arrayKey.count; $i++){
    $tempObj | Add-Member -MemberType NoteProperty -Name $arrayKey[$i] -Value $arrayValue[$i]    
  }
 $arrayOfInfo.Add($tempObj) | out-null 

}

$moreStuff = "Hello", "Penish"
$otherStuff = "Yes", "no"

createList -arrayKey $moreStuff -arrayValue $otherStuff
$PCName = HOSTNAME
$userInfo = "$env:username", "$env:USERDOMAIN", "$PCName"


$listOfInfo = "Username", "Domain", "PCName"

createList -arrayKey $listOfInfo -arrayValue $userInfo
$arrayOfInfo | format-list
#$ranges | format-list | out-file "C:\users\josh\desktop\test.csv" 

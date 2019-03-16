$validAdmins = [System.Collections.ArrayList]@()
Get-LocalGroupMember Administrators | Select-Object SID | ForEach-Object { 
   $tempObject = New-Object System.Object
   $tempObject = Get-LocalUser $_.SID | Where-Object {$_.Enabled -eq $true } | Select-Object Name, Enabled
   $validAdmins.add($tempObject) | Out-Null
}
$validAdmins

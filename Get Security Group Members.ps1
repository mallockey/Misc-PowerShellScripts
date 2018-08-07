import-module ActiveDirectory
$ou = "OU=Groups,OU=New York,DC=,DC=corp"
$grouparray = get-childitem -path AD:\$ou | select -expandproperty name

  foreach ($group in $grouparray)
  {
    if($group -like "*\*")
    {
    $group = $group.replace("\","_")
    }

  write-output "Group name:" $group | out-file "C:\users\eciadmin\desktop\securitygroups.csv" -append
  get-adgroupmember -identity $group | select name | sort-object name | out-file "C:\users\eciadmin\desktop\securitygroups.csv" -append
  write-output "==================================================" | out-file "C:\users\eciadmin\desktop\securitygroups.csv" -append
  }
read-host "Completed"

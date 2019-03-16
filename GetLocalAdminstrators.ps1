$localValidAdmins = Get-LocalGroupMember administrators | select sid | ForEach-Object {
Get-LocalUser $_.sid | where {$_.Enabled -eq $true } | select name, enabled}

$localValidAdmins

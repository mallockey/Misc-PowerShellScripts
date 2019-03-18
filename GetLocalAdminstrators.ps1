$validLocalAdmins = Get-LocalGroupMember administrators | where-object {$_.PrincipalSource -ne "ActiveDirectory"} | select sid | ForEach-Object {
Get-LocalUser $_.sid | where {$_.Enabled -eq $true } | select name, enabled
}

$validLocalAdmins

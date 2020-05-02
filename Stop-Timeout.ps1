$wshell = New-Object -ComObject wscript.shell
while($True){
	Start-Sleep -Seconds 1
	$wshell.SendKeys('{CAPSLOCK}')
}

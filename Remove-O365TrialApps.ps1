<#
Sample Data:
"scenario=install scenariosubtype=ARP sourcetype=None productstoremove=O365HomePremRetail.16_fr-fr_x-none culture=fr-fr DisplayLevel=False"
#>

$AllLanguages =  "en-us",
                 "es-es",
                 "fr-fr"

$ClickToRunPath = "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe" 
foreach($Language in $AllLanguages){
    Start-Process $ClickToRunPath -ArgumentList "scenario=install scenariosubtype=ARP sourcetype=None productstoremove=O365HomePremRetail.16_$($Language)_x-none culture=$($Language) DisplayLevel=False" -Wait
    Start-Sleep -Seconds 5
}

Param(
    [Parameter(Mandatory=$true)]
    $Properties
)
Write-Host -ForegroundColor yellow "INFO: If Property does not exist for PC, it will not be outputted." 
$properties = $properties.Split(" ")
import-module ActiveDirectory
$currentDir = "$psscriptroot"
$computersArray = get-adcomputer -filter * | select -expandproperty name
$PCInfo = [System.Collections.ArrayList]@()

foreach($computer in $computersArray){
    write-progress -Activity "Collecting Data" -Status "Current PC: $computer"
    $ADInfo = get-adcomputer -identity $computer -properties *
    $PCObj = New-Object -TypeName PSObject 
    $PCObj | Add-Member -MemberType NoteProperty -Name "PCName" -Value $computer
    foreach($prop in $properties) { 
        if(!$ADInfo[$prop]){
            continue
        }
        $currentProp = $ADInfo | select -ExpandProperty $prop #only did it this way to remove curly braces from output
        $PCObj  | Add-Member -MemberType NoteProperty -Name $prop -Value $currentProp
    }
     $PCInfo.Add($PCObj) | out-null
}
Write-Host -------------------------------------------
$PCInfo | format-list
Write-Host  -------------------------------------------
Write-Host "ComputerProperties.csv was exported to $currentDir" -ForegroundColor yellow
$PCInfo  | export-csv $currentDir\"ComputerProperties.csv" -noTypeInformation

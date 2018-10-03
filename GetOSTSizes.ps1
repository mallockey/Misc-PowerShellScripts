$computersArray = get-content computers.txt
foreach($computer in $computersArray){
 
    $ostArray = Get-ChildItem -Path "\\$computer\c$\users\*\appdata\local\microsoft\outlook\*.ost
    write-host "============================================================"
    write-host $computer
    write-host "============================================================"
        foreach($ost in $ostArray){
        $ostLength = $ost.length
        $ostLengthInGB = $ostLength / 1gb
        $ostLengthInGb = [int]$ostLengthInGb
        $ostTotalSum += $ost.length
        $ostTotalSumInGb = $ostTotalSum / 1gb
        $ostTotalSumInGb = [int]$ostTotalSumInGb

        write-host $ost.name "|" $ostLengthInGB
        }
        $ostTotalSum = $ostTotalSum / 1gb
        $ostTotalSum = [int]$ostTotalSum
        write-host Total Sum: $ostTotalSum "GBs"
        
write-host "============================================================"
write-host "End of "$computer
write-host "============================================================"
}
read-host "wait"

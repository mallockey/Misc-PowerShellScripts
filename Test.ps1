

 $CSV = Import-Csv -LiteralPath .\test.csv | where-object {$_.Location -eq "New York"}

 write-host $csv.PCName


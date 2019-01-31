$subDirectory = "C:\Users\Jmelo\desktop\test"
$monthsToNumbers = @{
    "01" = "January"
    "02" = "February"
    "03" = "March"
    "04" = "April"
    "05" = "May"
    "06" = "June"
    "07" = "July"
    "08" = "August"
    "09" = "September"
    "10" = "October"
    "11" = "November"
    "12" = "December"
}

function getCurrentMonthFolder{
    Param(
    $date
    )	
$todaysMonth = $date.Substring(0,2)
$month = $monthsToNumbers[$todaysMonth]
$year = $date.Substring(6)
$testIfYearExists = test-path -path $subDirectory\$year	
    if($testIfYearExists -eq $false){
    $makeNewYearDir = New-Item -Type directory "$subDirectory\$year"
    }
$testIfMonthExists = test-path -path $subDirectory\$year\$month
    if($testIfMonthExists -eq $false){
    $makeNewMonthDir =  New-Item -Type directory "$subDirectory\$year\$month"
    }
$finalDirectory = "$subDirectory\$year\$month"
return $finalDirectory
}
function getDate{  
$todaysDate = get-date -UFormat %D
$todaysDate = $todaysDate.Replace("/","-")
$todaysDate = $todaysDate.Insert(6,"20")
return $todaysDate
}
$currentDate = getDate
$finalDirectory = getCurrentMonthFolder $currentDate
write-output $finalDirectory

#Do stuff below with $finalDirectory

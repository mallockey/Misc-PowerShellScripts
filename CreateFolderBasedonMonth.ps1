$subDirectory = "C:\Users\Josh\desktop"

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

$testIfMonthExists = test-path -path $subDirectory\$month
	if($testIfMonthExists -eq $false){
	$makeNewDir =  New-Item -Type directory "$subDirectory\$month"
	}
	return $month
}

function getDate{
    
    $todaysDate = get-date -UFormat %D
    $todaysDate = $todaysDate.Replace("/","-")
    
    return $todaysDate

}
$currentDate = getDate
$currentMonth = getCurrentMonthFolder $currentDate
write-output $currentDate
write-output $currentMonth

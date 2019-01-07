function getSizeOfFolder{

    Param(
    $directory
    )
	try{

    }
	catch{
        write-host "No permissions to folder "
		$total+= 0
    }
   
    if($total -gt 1000000000){  
        $total =  [math]::Round($total / 1gb, 2)
        return [String]$total += "GB"
    }
        elseif($total -gt 1000000){
        $total =  [math]::Round($total / 1mb,2)
        return [String]$total += "MB"
    }
    else{
        write-host 
        return [String]$total = "<1MB"
	}
	
    
}
$total = get-childitem "C:\Windows\System32"-recurse -erroraction SilentlyContinue| measure-object -Property Length -sum -erroraction stop | select -expandproperty sum
write-host $total
foreach($folder in $total){
    write-host $folder
  getSizeOfFolder $folder
}

read-host

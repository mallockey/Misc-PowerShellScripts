import-module activedirectory
write-host "================================================="
write-host "1) Check for PCs in AD, one by one"
write-host "2) Check for PCs in AD, by range"
write-host "================================================="
while($decision -ne "exit"){
$decision = read-host "Please make a decision"

    switch($decision){
        1 { 
            while($computer -ne "quit"){
                try{
                $computer = read-host "Enter PC name to be searched in AD"

                $AdPC = get-adcomputer -identity $computer -erroraction stop | select -expandproperty name
                $OU = get-adcomputer -identity $computer -erroraction stop | select -expandproperty distinguishedname

                write-host Computer: $ADPC "is in Active Directory"
                write-host OU: $ou

                }
                catch{
                    write-host $computer "is not in AD"
                }
            }
          write-host "================================================="
        }
    2{  
        [int]$endRange = 0
        $startPC = read-host "Enter PC name convention(Ex. WS-0)"
        [int]$startRange = read-host "Enter first number"
        $endRange = read-host "Enter end range"
        for($i=$startRange; $i -le $endRange; $i++ ){
	        try{  
             $AdPC = get-adcomputer -identity $startPC$i -erroraction stop | select -expandproperty name
             $OU = get-adcomputer -identity $startPC$i -erroraction stop | select -expandproperty distinguishedname

	          write-host $AdPC "is in Active Directory"
	          write-host OU: $OU
	          write-host "================================================="
            }
            catch{
            write-host "$startPC$Ii" "is not in AD"
            }
        }
          write-host "================================================="
    }
    3{
        write-host "Exiting"

    }

}
}
read-host "Press Enter to exit"

<#
Author: Josh Melo
Last Updated: 11/26/18
This script will copy a single file to all the PCs in the Computers.txt file.
#>
try{
$computersArray = get-content computers.txt
}
catch{
write-host "Computers.txt not found"
}

$WorkSpaceApp = ""
$app = split-path -leaf $WorkSpaceApp
foreach($computers in $computersArray){
    if(test-connection -computerName $computers -count 1 -quiet){
    $computersAndLink = "\\"+$computers+"\c$\kits"
    try{
    copy-item -path $WorkSpaceApp -destination $computersAndLink
    write-host -foreGroundcolor green "Copied $app to $computersAndLink"
    }
    catch{
    write-host -foregroundColor red $computer "copy did not complete"
    }	
    }
    else{
    write-host -foregroundColor red $computers "is not online."
    }
}


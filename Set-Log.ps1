function Get-TimeStamp{
    return ((Get-Date -format d)) + " " + (Get-Date -Format t)
}
function Add-ToLog{

    param(
        [Switch]$Info,
        [Switch]$GeneratedErrorMessage,
        [Switch]$CustomErrorMessage,
        [String]$Message,
        [System.Collections.ArrayList]$LogArray,
        [Switch]$Exit
    )
    if($Info){
        $LogArray.Add("$(Get-TimeStamp) Info: $Message") | Out-Null
    }elseif($CustomErrorMessage){
        $LogArray.Add("$(Get-TimeStamp) Error: $Message") | Out-Null
    }elseif($GeneratedErrorMessage){
        $LogArray.Add("$(Get-TimeStamp) GeneratedErrorMessage: $Message") | Out-Null
    }

    if($Exit){
        $LogArray | Out-File $LogFile
        exit
    }

}

$CurrentLog = [System.Collections.ArrayList]@()
$LogFile = "C:\Users\Josh\Desktop\Log.txt"
Add-ToLog -LogArray $CurrentLog -Message "Hello there!" -Info -Exit

function Parse-ShadowOutput {

    param(
        $StringToParse
    )

    $IndexOfColon = $StringToParse.IndexOf(":")
    return $StringToParse.SubString($IndexOfColon + 2)

}

$ShadowProps = [Ordered]@{
    DriveLetter = $null
    UsedSpace = $null
    AllocatedSpace = $null
    MaxSpace = $null
}

$ArrayOfShadowInfo = [System.Collections.ArrayList]@()
$ShadowCopyInfo = vssadmin list shadowstorage

for($i = 0; $i -lt $ShadowCopyInfo.Length; $i++){
    if($ShadowCopyInfo[$i] -like "*For volume*"){
        $ShadowObj = New-Object -TypeName PSObject -Property $ShadowProps

        $IndexOfParen = $ShadowCopyInfo[$i].IndexOf(":")
        $ShadowObj.DriveLetter = $ShadowCopyInfo[$i].SubString($IndexOfParen + 3, 1)    #DriveLetter
      
        $ShadowObj.UsedSpace = Parse-ShadowOutput -StringToParse $ShadowCopyInfo[$i + 2]
        $ShadowObj.AllocatedSpace = Parse-ShadowOutput -StringToParse $ShadowCopyInfo[$i + 3]
        $ShadowObj.MaxSpace = Parse-ShadowOutput -StringToParse $ShadowCopyInfo[$i + 4]
        $ArrayOfShadowInfo.Add($ShadowObj) | Out-Null
    }
}

$ArrayOfShadowInfo

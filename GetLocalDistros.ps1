$olFolderContacts = 10
$outlook = New-Object -ComObject Outlook.Application
$contacts = $outlook.session.GetDefaultFolder(10).items
$namespace = $Outlook.GetNameSpace("mapi")
$folders = $namespace.getDefaultFolder($olFolderContacts)
$arrayOfDistros = [System.Collections.ArrayList]@()
$distroGroups = $folders | ForEach-Object {$_.Items} | 
                Where-Object {$_.DLName -ne $null}

foreach($distro in $distroGroups){
    $tempObj = New-Object -TypeName PSObject 
    $tempObj | Add-Member -MemberType NoteProperty -Name DLName -Value $distro.DLName
    for($i=1; $i -le $distro.MemberCount; $i++){
        $currentDistroMember = $distro.GetMember($i) | Select-Object -ExpandProperty address
        $tempObj | Add-Member -MemberType NoteProperty -Name $i -Value $currentDistroMember
    }
    $arrayOfDistros.Add($tempObj) | Out-Null
}   
$arrayOfDistros | Format-List
$arrayOfDistros | Format-List |  Out-File "C:\Kits\LocalDistros.txt" 
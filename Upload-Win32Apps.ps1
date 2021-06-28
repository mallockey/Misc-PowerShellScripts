<#
.COPYRIGHT
Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
See LICENSE in the project root for license information.
#>
param(
    [ValidateSet("AdobeAcrobatReaderDC", "AdobeAcrobatDCPro", "AdobeCreativeCloud-User", "AutomateAgent",  "Bloomberg", "CitrixWorkspace",
                 "CiscoAnyConnect", "CiscoWebexPlayer-ARF","CiscoWebexPlayer-WRF", "Dropbox", "GoogleChrome", "LogitechUnifyingSoftware", 
                 "MozillaFirefox","Mimecast","SEPSmallBusiness","SevenZip", "WebExMeetings","Zoom")]
                 [String[]]$AppsToUpload,
				 [Switch]$AllApps,
                 [Switch]$AssignApps,
                 [Int]$LabtechLocationID
) 
function CloneObject($object) {

    $stream = New-Object IO.MemoryStream;
    $formatter = New-Object Runtime.Serialization.Formatters.Binary.BinaryFormatter;
    $formatter.Serialize($stream, $object);
    $stream.Position = 0;
    $formatter.Deserialize($stream);
}
function WriteHeaders($authToken) {

    foreach ($header in $authToken.GetEnumerator()) {
        if ($header.Name.ToLower() -eq "authorization") {
            continue;
        }

        Write-Host -ForegroundColor Gray "$($header.Name): $($header.Value)";
    }
}
function MakeGetRequest($collectionPath) {

    $uri = "$baseUrl$collectionPath";
    $request = "GET $uri";
	
    if ($logRequestUris) { Write-Host $request; }
    if ($logHeaders) { WriteHeaders $authToken; }

    try {
        $response = Invoke-RestMethod $uri -Method Get -Headers $authToken;
        $response;
    }
    catch {
        Write-Host -ForegroundColor Red $request;
        Write-Host -ForegroundColor Red $_.Exception.Message;
        throw;
    }
}
function MakePatchRequest($collectionPath, $body) {
    MakeRequest "PATCH" $collectionPath $body;
}
function MakePostRequest($collectionPath, $body) {
    MakeRequest "POST" $collectionPath $body;
}
function MakeRequest($verb, $collectionPath, $body) {

    $uri = "$baseUrl$collectionPath";
    $request = "$verb $uri";
	
    $clonedHeaders = CloneObject $authToken;
    $clonedHeaders["content-length"] = $body.Length;
    $clonedHeaders["content-type"] = "application/json";

    if ($logRequestUris) { Write-Host $request; }
    if ($logHeaders) { WriteHeaders $clonedHeaders; }
    if ($logContent) { Write-Host -ForegroundColor Gray $body; }

    try {
        $response = Invoke-RestMethod $uri -Method $verb -Headers $clonedHeaders -Body $body;
        $response;
    }
    catch {
        Write-Host -ForegroundColor Red $request;
        Write-Host -ForegroundColor Red $_.Exception.Message;
        throw;
    }
}
function UploadAzureStorageChunk($sasUri, $id, $body) {

    $uri = "$sasUri&comp=block&blockid=$id";
    $request = "PUT $uri";

    $iso = [System.Text.Encoding]::GetEncoding("iso-8859-1");
    $encodedBody = $iso.GetString($body);
    $headers = @{
        "x-ms-blob-type" = "BlockBlob"
    };

    if ($logRequestUris) { Write-Host $request; }
    if ($logHeaders) { WriteHeaders $headers; }

    try {
        $response = Invoke-WebRequest $uri -Method Put -Headers $headers -Body $encodedBody;
    }
    catch {
        Write-Host -ForegroundColor Red $request;
        Write-Host -ForegroundColor Red $_.Exception.Message;
        throw;
    }

}
function FinalizeAzureStorageUpload($sasUri, $ids) {

    $uri = "$sasUri&comp=blocklist";
    $request = "PUT $uri";

    $xml = '<?xml version="1.0" encoding="utf-8"?><BlockList>';
    foreach ($id in $ids) {
        $xml += "<Latest>$id</Latest>";
    }
    $xml += '</BlockList>';

    if ($logRequestUris) { Write-Host $request; }
    if ($logContent) { Write-Host -ForegroundColor Gray $xml; }

    try {
        Invoke-RestMethod $uri -Method Put -Body $xml;
    }
    catch {
        Write-Host -ForegroundColor Red $request;
        Write-Host -ForegroundColor Red $_.Exception.Message;
        throw;
    }
}
function UploadFileToAzureStorage($sasUri, $filepath, $fileUri) {

    try {

        $chunkSizeInBytes = 1024l * 1024l * $azureStorageUploadChunkSizeInMb;
		
        # Start the timer for SAS URI renewal.
        $sasRenewalTimer = [System.Diagnostics.Stopwatch]::StartNew()
		
        # Find the file size and open the file.
        $fileSize = (Get-Item $filepath).length;
        $chunks = [Math]::Ceiling($fileSize / $chunkSizeInBytes);
        $reader = New-Object System.IO.BinaryReader([System.IO.File]::Open($filepath, [System.IO.FileMode]::Open));
        $position = $reader.BaseStream.Seek(0, [System.IO.SeekOrigin]::Begin);
		
        # Upload each chunk. Check whether a SAS URI renewal is required after each chunk is uploaded and renew if needed.
        $ids = @();

        for ($chunk = 0; $chunk -lt $chunks; $chunk++) {

            $id = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($chunk.ToString("0000")));
            $ids += $id;

            $start = $chunk * $chunkSizeInBytes;
            $length = [Math]::Min($chunkSizeInBytes, $fileSize - $start);
            $bytes = $reader.ReadBytes($length);
			
            $currentChunk = $chunk + 1;			

            Write-Progress -Activity "Uploading File to Azure Storage" -status "Uploading chunk $currentChunk of $chunks" `
                -percentComplete ($currentChunk / $chunks*100)

            $uploadResponse = UploadAzureStorageChunk $sasUri $id $bytes;
			
            # Renew the SAS URI if 7 minutes have elapsed since the upload started or was renewed last.
            if ($currentChunk -lt $chunks -and $sasRenewalTimer.ElapsedMilliseconds -ge 450000) {

                $renewalResponse = RenewAzureStorageUpload $fileUri;
                $sasRenewalTimer.Restart();
			
            }

        }

        Write-Progress -Completed -Activity "Uploading File to Azure Storage"

        $reader.Close();

    }

    finally {

        if ($reader -ne $null) { $reader.Dispose(); }
	
    }
	
    # Finalize the upload.
    $uploadResponse = FinalizeAzureStorageUpload $sasUri $ids;

}
function RenewAzureStorageUpload($fileUri) {

    $renewalUri = "$fileUri/renewUpload";
    $actionBody = "";
    $rewnewUriResult = MakePostRequest $renewalUri $actionBody;
	
    $file = WaitForFileProcessing $fileUri "AzureStorageUriRenewal" $azureStorageRenewSasUriBackOffTimeInSeconds;

}
function WaitForFileProcessing($fileUri, $stage) {

    $attempts = 600;
    $waitTimeInSeconds = 10;

    $successState = "$($stage)Success";
    $pendingState = "$($stage)Pending";
    $failedState = "$($stage)Failed";
    $timedOutState = "$($stage)TimedOut";

    $file = $null;
    while ($attempts -gt 0) {
        $file = MakeGetRequest $fileUri;

        if ($file.uploadState -eq $successState) {
            break;
        }
        elseif ($file.uploadState -ne $pendingState) {
            Write-Host -ForegroundColor Red $_.Exception.Message;
            throw "File upload state is not success: $($file.uploadState)";
        }

        Start-Sleep $waitTimeInSeconds;
        $attempts--;
    }

    if ($file -eq $null -or $file.uploadState -ne $successState) {
        throw "File request did not complete in the allotted time.";
    }

    $file;
}
function GetWin32AppBody() {

    param
    (

        [parameter(Mandatory = $true, ParameterSetName = "MSI", Position = 1)]
        [Switch]$MSI,

        [parameter(Mandatory = $true, ParameterSetName = "EXE", Position = 1)]
        [Switch]$EXE,

        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$displayName,

        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$publisher,

        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$description,

        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$filename,

        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$SetupFileName,

        [parameter(Mandatory = $true)]
        [ValidateSet('system', 'user')]
        $installExperience = "system",

        [parameter(Mandatory = $true, ParameterSetName = "EXE")]
        [ValidateNotNullOrEmpty()]
        $installCommandLine,

        [parameter(Mandatory = $true, ParameterSetName = "EXE")]
        [ValidateNotNullOrEmpty()]
        $uninstallCommandLine,

        [parameter(Mandatory = $true, ParameterSetName = "MSI")]
        [ValidateNotNullOrEmpty()]
        $MsiPackageType,

        [parameter(Mandatory = $true, ParameterSetName = "MSI")]
        [ValidateNotNullOrEmpty()]
        $MsiProductCode,

        [parameter(Mandatory = $false, ParameterSetName = "MSI")]
        $MsiProductName,

        [parameter(Mandatory = $true, ParameterSetName = "MSI")]
        [ValidateNotNullOrEmpty()]
        $MsiProductVersion,

        [parameter(Mandatory = $false, ParameterSetName = "MSI")]
        $MsiPublisher,

        [parameter(Mandatory = $true, ParameterSetName = "MSI")]
        [ValidateNotNullOrEmpty()]
        $MsiRequiresReboot,

        [parameter(Mandatory = $true, ParameterSetName = "MSI")]
        [ValidateNotNullOrEmpty()]
        $MsiUpgradeCode

    )

    if ($MSI) {

        $body = @{ "@odata.type" = "#microsoft.graph.win32LobApp" };
        $body.applicableArchitectures = "x64";
        $body.description = $description;
        $body.developer = "";
        $body.displayName = $displayName;
        $body.fileName = $filename;
        $body.installCommandLine = "msiexec /i `"$SetupFileName`""
        $body.installExperience = @{"runAsAccount" = "$installExperience" };
        $body.informationUrl = $null;
        $body.isFeatured = $false;
        $body.minimumSupportedOperatingSystem = @{"v10_1903" = $true };
        $body.msiInformation = @{
            "packageType"    = "$MsiPackageType";
            "productCode"    = "$MsiProductCode";
            "productName"    = "$MsiProductName";
            "productVersion" = "$MsiProductVersion";
            "publisher"      = "$MsiPublisher";
            "requiresReboot" = "$MsiRequiresReboot";
            "upgradeCode"    = "$MsiUpgradeCode"
        };
        $body.notes = "";
        $body.owner = "";
        $body.privacyInformationUrl = $null;
        $body.publisher = $publisher;
        $body.runAs32bit = $false;
        $body.setupFilePath = $SetupFileName;
        $body.uninstallCommandLine = "msiexec /x `"$MsiProductCode`""

    }

    elseif ($EXE) {

        $body = @{ "@odata.type" = "#microsoft.graph.win32LobApp" };
        $body.applicableArchitectures = "x64";
        $body.description = $description;
        $body.developer = "";
        $body.displayName = $displayName;
        $body.fileName = $filename;
        $body.installCommandLine = "$installCommandLine"
        $body.installExperience = @{"runAsAccount" = "$installExperience" };
        $body.informationUrl = $null;
        $body.isFeatured = $false;
        $body.minimumSupportedOperatingSystem = @{"v10_1903" = $true };
        $body.msiInformation = $null;
        $body.notes = "";
        $body.owner = "";
        $body.privacyInformationUrl = $null;
        $body.publisher = $publisher;
        $body.runAs32bit = $false;
        $body.setupFilePath = $SetupFileName;
        $body.uninstallCommandLine = "$uninstallCommandLine"

    }

    $body;
}
function GetAppFileBody($name, $size, $sizeEncrypted, $manifest) {

    $body = @{ "@odata.type" = "#microsoft.graph.mobileAppContentFile" };
    $body.name = $name;
    $body.size = $size;
    $body.sizeEncrypted = $sizeEncrypted;
    $body.manifest = $manifest;
    $body.isDependency = $false;

    $body;
}
function GetAppCommitBody($contentVersionId, $LobType) {

    $body = @{ "@odata.type" = "#$LobType" };
    $body.committedContentVersion = $contentVersionId;

    $body;

}
function Test-SourceFile() {
    param(
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        $SourceFile
    )
    try {
        if (!(test-path "$SourceFile")) {
            Write-Host
            Write-Host "Source File '$sourceFile' doesn't exist..." -ForegroundColor Red
            throw
        }
    }
    catch {
        Write-Host -ForegroundColor Red $_.Exception.Message;
        Write-Host
        break
    }
}
function New-DetectionRule() {

    [cmdletbinding()]

    param
    (
        [parameter(Mandatory = $true, ParameterSetName = "PowerShell", Position = 1)]
        [Switch]$PowerShell,

        [parameter(Mandatory = $true, ParameterSetName = "MSI", Position = 1)]
        [Switch]$MSI,

        [parameter(Mandatory = $true, ParameterSetName = "File", Position = 1)]
        [Switch]$File,

        [parameter(Mandatory = $true, ParameterSetName = "Registry", Position = 1)]
        [Switch]$Registry,

        [parameter(Mandatory = $true, ParameterSetName = "PowerShell")]
        [ValidateNotNullOrEmpty()]
        [String]$ScriptFile,

        [parameter(Mandatory = $true, ParameterSetName = "PowerShell")]
        [ValidateNotNullOrEmpty()]
        $enforceSignatureCheck,

        [parameter(Mandatory = $true, ParameterSetName = "PowerShell")]
        [ValidateNotNullOrEmpty()]
        $runAs32Bit,

        [parameter(Mandatory = $true, ParameterSetName = "MSI")]
        [ValidateNotNullOrEmpty()]
        [String]$MSIproductCode,
   
        [parameter(Mandatory = $true, ParameterSetName = "File")]
        [ValidateNotNullOrEmpty()]
        [String]$Path,
 
        [parameter(Mandatory = $true, ParameterSetName = "File")]
        [ValidateNotNullOrEmpty()]
        [string]$FileOrFolderName,

        [parameter(Mandatory = $true, ParameterSetName = "File")]
        [ValidateSet("notConfigured", "exists", "modifiedDate", "createdDate", "version", "sizeInMB")]
        [string]$FileDetectionType,

        [parameter(Mandatory = $false, ParameterSetName = "File")]
        $FileDetectionValue = $null,

        [parameter(Mandatory = $true, ParameterSetName = "File")]
        [ValidateSet("True", "False")]
        [string]$check32BitOn64System = "False",

        [parameter(Mandatory = $true, ParameterSetName = "Registry")]
        [ValidateNotNullOrEmpty()]
        [String]$RegistryKeyPath,

        [parameter(Mandatory = $true, ParameterSetName = "Registry")]
        [ValidateSet("notConfigured", "exists", "doesNotExist", "string", "integer", "version")]
        [string]$RegistryDetectionType,

        [parameter(Mandatory = $false, ParameterSetName = "Registry")]
        [ValidateNotNullOrEmpty()]
        [String]$RegistryValue,

        [parameter(Mandatory = $true, ParameterSetName = "Registry")]
        [ValidateSet("True", "False")]
        [string]$check32BitRegOn64System = "False"

    )

    if ($PowerShell) {

        if (!(Test-Path "$ScriptFile")) {
            
            Write-Host
            Write-Host "Could not find file '$ScriptFile'..." -ForegroundColor Red
            Write-Host "Script can't continue..." -ForegroundColor Red
            Write-Host
            break

        }
        
        $ScriptContent = [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes("$ScriptFile"));
        
        $DR = @{ "@odata.type" = "#microsoft.graph.win32LobAppPowerShellScriptDetection" }
        $DR.enforceSignatureCheck = $false;
        $DR.runAs32Bit = $false;
        $DR.scriptContent = "$ScriptContent";

    }
    
    elseif ($MSI) {
    
        $DR = @{ "@odata.type" = "#microsoft.graph.win32LobAppProductCodeDetection" }
        $DR.productVersionOperator = "notConfigured";
        $DR.productCode = "$MsiProductCode";
        $DR.productVersion = $null;

    }

    elseif ($File) {
    
        $DR = @{ "@odata.type" = "#microsoft.graph.win32LobAppFileSystemDetection" }
        $DR.check32BitOn64System = "$check32BitOn64System";
        $DR.detectionType = "$FileDetectionType";
        $DR.detectionValue = $FileDetectionValue;
        $DR.fileOrFolderName = "$FileOrFolderName";
        $DR.operator = "notConfigured";
        $DR.path = "$Path"

    }

    elseif ($Registry) {
    
        $DR = @{ "@odata.type" = "#microsoft.graph.win32LobAppRegistryDetection" }
        $DR.check32BitOn64System = "$check32BitRegOn64System";
        $DR.detectionType = "$RegistryDetectionType";
        $DR.detectionValue = "";
        $DR.keyPath = "$RegistryKeyPath";
        $DR.operator = "notConfigured";
        $DR.valueName = "$RegistryValue"

    }

    return $DR

}
function Get-DefaultReturnCodes() {

    @{"returnCode" = 0; "type" = "success" }, `
    @{"returnCode" = 1707; "type" = "success" }, `
    @{"returnCode" = 3010; "type" = "softReboot" }, `
    @{"returnCode" = 1641; "type" = "hardReboot" }, `
    @{"returnCode" = 1618; "type" = "retry" }

}
function New-ReturnCode() {

    param
    (
        [parameter(Mandatory = $true)]
        [int]$returnCode,
        [parameter(Mandatory = $true)]
        [ValidateSet('success', 'softReboot', 'hardReboot', 'retry')]
        $type
    )

    @{"returnCode" = $returnCode; "type" = "$type" }

}
function Get-IntuneWinXML() {

    param
    (
        [Parameter(Mandatory = $true)]
        $SourceFile,

        [Parameter(Mandatory = $true)]
        $fileName,

        [Parameter(Mandatory = $false)]
        [ValidateSet("false", "true")]
        [string]$removeitem = "true"
    )

    Test-SourceFile "$SourceFile"

    $Directory = [System.IO.Path]::GetDirectoryName("$SourceFile")

    Add-Type -Assembly System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead("$SourceFile")

    $zip.Entries | Where-Object { $_.Name -like "$filename" } | ForEach-Object {

        [System.IO.Compression.ZipFileExtensions]::ExtractToFile($_, "$Directory\$filename", $true)

    }

    $zip.Dispose()

    [xml]$IntuneWinXML = Get-Content "$Directory\$filename"

    return $IntuneWinXML

    if ($removeitem -eq "true") { remove-item "$Directory\$filename" }

}
function Get-IntuneWinFile() {

    param
    (
        [Parameter(Mandatory = $true)]
        $SourceFile,

        [Parameter(Mandatory = $true)]
        $fileName,

        [Parameter(Mandatory = $false)]
        [string]$Folder = "win32"
    )

    $Directory = [System.IO.Path]::GetDirectoryName("$SourceFile")

    if (!(Test-Path "$Directory\$folder")) {

        New-Item -ItemType Directory -Path "$Directory" -Name "$folder" | Out-Null

    }

    Add-Type -Assembly System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead("$SourceFile")

    $zip.Entries | where { $_.Name -like "$filename" } | foreach {

        [System.IO.Compression.ZipFileExtensions]::ExtractToFile($_, "$Directory\$folder\$filename", $true)

    }

    $zip.Dispose()

    return "$Directory\$folder\$filename"

}
function Upload-Win32Lob() {

    <#
.SYNOPSIS
This function is used to upload a Win32 Application to the Intune Service
.DESCRIPTION
This function is used to upload a Win32 Application to the Intune Service
.EXAMPLE
Upload-Win32Lob "C:\Packages\package.intunewin" -publisher "Microsoft" -description "Package"
This example uses all parameters required to add an intunewin File into the Intune Service
.NOTES
NAME: Upload-Win32LOB
#>

    [cmdletbinding()]

    param
    (
        [parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string]$SourceFile,

        [parameter(Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        [string]$displayName,

        [parameter(Mandatory = $true, Position = 2)]
        [ValidateNotNullOrEmpty()]
        [string]$publisher,

        [parameter(Mandatory = $true, Position = 3)]
        [ValidateNotNullOrEmpty()]
        [string]$description,

        [parameter(Mandatory = $true, Position = 4)]
        [ValidateNotNullOrEmpty()]
        $detectionRules,

        [parameter(Mandatory = $true, Position = 5)]
        [ValidateNotNullOrEmpty()]
        $returnCodes,

        [parameter(Mandatory = $false, Position = 6)]
        [ValidateNotNullOrEmpty()]
        [string]$installCmdLine,

        [parameter(Mandatory = $false, Position = 7)]
        [ValidateNotNullOrEmpty()]
        [string]$uninstallCmdLine,

        [parameter(Mandatory = $false, Position = 8)]
        [ValidateSet('system', 'user')]
        $installExperience = "system"
    )

    try	{

        $LOBType = "microsoft.graph.win32LobApp"

        Write-Host "Testing if SourceFile '$SourceFile' Path is valid..." -ForegroundColor Yellow
  
        Write-Host "Creating JSON data to pass to the service..." -ForegroundColor Yellow

        # Funciton to read Win32LOB file
        $DetectionXML = Get-IntuneWinXML "$SourceFile" -fileName "detection.xml"
        
        $FileName = $DetectionXML.ApplicationInfo.FileName
        $SetupFileName = $DetectionXML.ApplicationInfo.SetupFile

        $Ext = [System.IO.Path]::GetExtension($SetupFileName)

        if ((($Ext).contains("msi") -or ($Ext).contains("Msi")) -and (!$installCmdLine -or !$uninstallCmdLine)) {

            # MSI
            $MsiExecutionContext = $DetectionXML.ApplicationInfo.MsiInfo.MsiExecutionContext
            $MsiPackageType = "DualPurpose";
            if ($MsiExecutionContext -eq "System") { $MsiPackageType = "PerMachine" }
            elseif ($MsiExecutionContext -eq "User") { $MsiPackageType = "PerUser" }

            $MsiProductCode = $DetectionXML.ApplicationInfo.MsiInfo.MsiProductCode
            $MsiProductVersion = $DetectionXML.ApplicationInfo.MsiInfo.MsiProductVersion
            $MsiPublisher = $DetectionXML.ApplicationInfo.MsiInfo.MsiPublisher
            $MsiRequiresReboot = $DetectionXML.ApplicationInfo.MsiInfo.MsiRequiresReboot
            $MsiUpgradeCode = $DetectionXML.ApplicationInfo.MsiInfo.MsiUpgradeCode
            
            if ($MsiRequiresReboot -eq "false") { $MsiRequiresReboot = $false }
            elseif ($MsiRequiresReboot -eq "true") { $MsiRequiresReboot = $true }

            $mobileAppBody = GetWin32AppBody `
                -MSI `
                -displayName "$DisplayName" `
                -publisher "$publisher" `
                -description $description `
                -filename $FileName `
                -SetupFileName "$SetupFileName" `
                -installExperience $installExperience `
                -MsiPackageType $MsiPackageType `
                -MsiProductCode $MsiProductCode `
                -MsiProductName $displayName `
                -MsiProductVersion $MsiProductVersion `
                -MsiPublisher $MsiPublisher `
                -MsiRequiresReboot $MsiRequiresReboot `
                 -MsiUpgradeCode $MsiUpgradeCode
        }
        else {
            $mobileAppBody = GetWin32AppBody -EXE -displayName "$DisplayName" -publisher "$publisher" `
                -description $description -filename $FileName -SetupFileName "$SetupFileName" `
                -installExperience $installExperience -installCommandLine $installCmdLine `
                -uninstallCommandLine $uninstallcmdline
        }
        if ($DetectionRules.'@odata.type' -contains "#microsoft.graph.win32LobAppPowerShellScriptDetection" -and @($DetectionRules).'@odata.type'.Count -gt 1) {
            Write-Warning "A Detection Rule can either be 'Manually configure detection rules' or 'Use a custom detection script'"
            Write-Warning "It can't include both..."
            Write-Host
            break
        }
        else {
            $mobileAppBody | Add-Member -MemberType NoteProperty -Name 'detectionRules' -Value $detectionRules
        }
        #ReturnCodes
        if ($returnCodes) {
            $mobileAppBody | Add-Member -MemberType NoteProperty -Name 'returnCodes' -Value @($returnCodes)
        }
        else {
            Write-Host
            Write-Warning "Intunewin file requires ReturnCodes to be specified"
            Write-Warning "If you want to use the default ReturnCode run 'Get-DefaultReturnCodes'"
            Write-Host
            break
        }
        Write-Host "Creating application in Intune..." -ForegroundColor Yellow
        $mobileApp = MakePostRequest "mobileApps" ($mobileAppBody | ConvertTo-Json);
        # Get the content version for the new app (this will always be 1 until the new app is committed).
        Write-Host "Creating Content Version in the service for the application..." -ForegroundColor Yellow
        $appId = $mobileApp.id;
        $contentVersionUri = "mobileApps/$appId/$LOBType/contentVersions";
        $contentVersion = MakePostRequest $contentVersionUri "{}";
        # Encrypt file and Get File Information
        Write-Host "Getting Encryption Information for '$SourceFile'..." -ForegroundColor Yellow

        $encryptionInfo = @{ };
        $encryptionInfo.encryptionKey = $DetectionXML.ApplicationInfo.EncryptionInfo.EncryptionKey
        $encryptionInfo.macKey = $DetectionXML.ApplicationInfo.EncryptionInfo.macKey
        $encryptionInfo.initializationVector = $DetectionXML.ApplicationInfo.EncryptionInfo.initializationVector
        $encryptionInfo.mac = $DetectionXML.ApplicationInfo.EncryptionInfo.mac
        $encryptionInfo.profileIdentifier = "ProfileVersion1";
        $encryptionInfo.fileDigest = $DetectionXML.ApplicationInfo.EncryptionInfo.fileDigest
        $encryptionInfo.fileDigestAlgorithm = $DetectionXML.ApplicationInfo.EncryptionInfo.fileDigestAlgorithm

        $fileEncryptionInfo = @{ };
        $fileEncryptionInfo.fileEncryptionInfo = $encryptionInfo;

        # Extracting encrypted file
        $IntuneWinFile = Get-IntuneWinFile "$SourceFile" -fileName "$filename"

        [int64]$Size = $DetectionXML.ApplicationInfo.UnencryptedContentSize
        $EncrySize = (Get-Item "$IntuneWinFile").Length

        # Create a new file for the app.
        Write-Host "Creating a new file entry in Azure for the upload..." -ForegroundColor Yellow
        $contentVersionId = $contentVersion.id;
        $fileBody = GetAppFileBody "$FileName" $Size $EncrySize $null;
        $filesUri = "mobileApps/$appId/$LOBType/contentVersions/$contentVersionId/files";
        $file = MakePostRequest $filesUri ($fileBody | ConvertTo-Json);
	
        # Wait for the service to process the new file request.
        Write-Host "Waiting for the file entry URI to be created..." -ForegroundColor Yellow
        $fileId = $file.id;
        $fileUri = "mobileApps/$appId/$LOBType/contentVersions/$contentVersionId/files/$fileId";
        $file = WaitForFileProcessing $fileUri "AzureStorageUriRequest";

        # Upload the content to Azure Storage.
        Write-Host
        Write-Host "Uploading file to Azure Storage..." -f Yellow

        $sasUri = $file.azureStorageUri;
        UploadFileToAzureStorage $file.azureStorageUri "$IntuneWinFile" $fileUri;

        # Need to Add removal of IntuneWin file
        $IntuneWinFolder = [System.IO.Path]::GetDirectoryName("$IntuneWinFile")
        Remove-Item "$IntuneWinFile" -Force

        # Commit the file.
        Write-Host "Committing the file into Azure Storage..." -ForegroundColor Yellow
        $commitFileUri = "mobileApps/$appId/$LOBType/contentVersions/$contentVersionId/files/$fileId/commit";
        MakePostRequest $commitFileUri ($fileEncryptionInfo | ConvertTo-Json);

        # Wait for the service to process the commit file request.
        Write-Host "Waiting for the service to process the commit file request..." -ForegroundColor Yellow
        $file = WaitForFileProcessing $fileUri "CommitFile";

        # Commit the app.
        Write-Host "Committing the file into Azure Storage..." -ForegroundColor Yellow
        $commitAppUri = "mobileApps/$appId";
        $commitAppBody = GetAppCommitBody $contentVersionId $LOBType;
        MakePatchRequest $commitAppUri ($commitAppBody | ConvertTo-Json);

        Write-Host "Sleeping for $sleep seconds to allow patch completion..." -f Magenta
        Start-Sleep $sleep
    }
    catch {
        Write-Host -ForegroundColor Red "Aborting with exception: $($_.Exception.ToString())";
    }
}

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$BaseUrl = "https://graph.microsoft.com/beta/deviceAppManagement/"

$LogRequestUris = $False;
$LogHeaders = $False;
$LogContent = $False;
$AzureStorageUploadChunkSizeInMb = 6l;
$Sleep = 30

###########################################Convert all the folders in the .\RawDownloads folder to IntuneWin file##################################################
$AllSetupFolders = (Get-ChildItem -Path ".\RawDownloads" -Directory).FullName
$ApplicationInfo = Import-Csv ".\ApplicationDetails.csv"

foreach($SetupFolder in $AllSetupFolders){
    
    $CurrentAppName = Split-Path -Path $SetupFolder -Leaf
    $CurrentAppInfo = $ApplicationInfo | Where-Object {$_.ID -eq $CurrentAppName}

    if($AppsToUpload -contains $CurrentAppName -or $AllApps){
        
        if($CurrentAppInfo.MultipleFiles){
            $SetupFile = Resolve-Path $CurrentAppInfo.MultipleFiles
        }elseif($CurrentAppInfo.IsMSI -eq "TRUE"){
            $SetupFile = (Get-ChildItem -Path $SetupFolder -Filter "*.msi").FullName
        }else{
            $SetupFile = (Get-ChildItem -Path $SetupFolder -Filter "*.exe").FullName
        }

        if(!($SetupFile)){
            Add-ToLog -Message "Cannot find setup file for $($CurrentAppName). Confirm the setup file exists in RawDownloads" -Info
            continue
        }

        $OutputFolder = ".\IntuneWinUpload\$CurrentAppName\"
        New-Path $OutputFolder

        .\IntuneWinAppUtil.exe -c $SetupFolder -s $SetupFile -o $OutputFolder
  
        $IntuneWinFile = $OutputFolder + (Split-Path -Path $SetupFile -Leaf).Replace(".msi",".intunewin")
        $IntuneWinFile = $IntuneWinFile.Replace(".exe",".intunewin")

        if(Test-Path -Path $IntuneWinFile){
            Add-ToLog -Message "Converted $($CurrentAppName) to intunewin file successfully" -Info
        }else{
            Add-ToLog -Message "Failed to convert $($CurrentAppName) to intunewin file. Confirm the setup file exists in RawDownloads" -CustomErro
        }

        Start-Sleep -Seconds 2
    }
}

###########################################End of Convert all the folders in the .\RawDownloads folder to IntuneWin file##################################################

###############################################Upload Intune win files to Azure####################################################################
$PathToIntuneWinFiles = (Get-ChildItem -Path ".\IntuneWinUpload" -Filter "*.intunewin" -Recurse).FullName

foreach($Path in $PathToIntuneWinFiles){
    try{

        $CurrentInstallCommand = ""
        $CurrentAppID = Split-Path -Path (Split-Path -Path $Path -Parent) -Leaf
        $CurrentAppInfo = $ApplicationInfo | Where-Object {$_.ID -eq $CurrentAppID}

        if($AppsToUpload -contains $CurrentAppID -or $AllApps){

            if($CurrentAppID -eq "AutomateAgent"){
                $CurrentInstallCommand = "msiexec /i `"Agent_Install.MSI`" /quiet /norestart LOCATION=$LabtechLocationID"
            }elseif($CurrentAppInfo.MultipleFiles -eq "TRUE"){
                $CurrentInstallCommand = $CurrentAppInfo.InstallParameter #Static this needs to be updated manually if it changes.
            }elseif($CurrentAppInfo.IsMSI -eq "TRUE"){
                $CurrentAppInstallFile = (Split-Path -Path $Path -Leaf).Replace(".intunewin",".msi")
                if($CurrentAppInfo.InstallParameter -eq "STANDARD"){
                    $CurrentInstallCommand = "msiexec /i $CurrentAppInstallFile /qn"
                }else{
                    $CurrentInstallCommand = "msiexec /i $CurrentAppInstallFile $($CurrentAppInfo.InstallParameter)"
                }
            }else{
                $CurrentAppInstallFile = (Split-Path -Path $Path -Leaf).Replace(".intunewin",".exe")
                $CurrentInstallCommand = "$($CurrentAppInstallFile) $($CurrentAppInfo.InstallParameter)"
            }

            if($CurrentAppInfo.DetectionRule -like "*(x86)*"){
                $CheckFor32Bit = $True
            }else{
                $CheckFor32Bit = $False
            }
        
            if($CurrentAppInfo.DetectionRule -eq "PSScript"){
                $PSIntuneAppDetectionPath = (Get-ChildItem -Path ".\RawDownloads\$CurrentAppID\*.ps1").FullName
                $CurrentDetectionRule = New-DetectionRule -PowerShell -ScriptFile $PSIntuneAppDetectionPath `
                                                          -EnforceSignatureCheck $False `
                                                          -RunAs32Bit $False
        
            }else{
                $CurrentDetectionRule = New-DetectionRule -File -Path (Split-Path -Path $CurrentAppInfo.DetectionRule -Parent) `
                                                          -FileOrFolderName (Split-Path -Path $CurrentAppInfo.DetectionRule -Leaf) `
                                                          -FileDetectionType exists `
                                                          -check32BitOn64System $CheckFor32Bit

            }
        
            $CurrentDetectionRules = @($CurrentDetectionRule)

            $CurrentReturnCodes = Get-DefaultReturnCodes
            $CurrentReturnCodes += New-ReturnCode -returnCode 302 -type softReboot
            $CurrentReturnCodes += New-ReturnCode -returnCode 145 -type hardReboot
        
            Upload-Win32Lob -displayName $CurrentAppInfo.DisplayName -SourceFile "$Path" -publisher $CurrentAppInfo.Publisher `
                            -description $CurrentAppInfo.Description -detectionRules $CurrentDetectionRules -returnCodes $CurrentReturnCodes `
                            -installCmdLine $CurrentInstallCommand -uninstallCmdLine $CurrentAppInfo.UninstallCommand

            if(Get-IntuneConfigurationID -EndpointURL "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps" -DisplayName $CurrentAppInfo.DisplayName){
                Add-ToLog -Message "Uploaded $($CurrentAppID) to Azure successfully" -Info
            }else{
                Add-ToLog -Message "The app failed to upload with an unknown error" -CustomErrorMessage
            }
			
			if($AssignApps){
                Start-Sleep -Seconds 5
                $AppConfigID = Get-IntuneConfigurationID -EndpointURL "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps" -DisplayName $CurrentAppInfo.DisplayName
    
                $APIntuneDevicesID = Get-IntuneConfigurationID -EndpointURL "https://graph.microsoft.com/beta/groups" -DisplayName "AP_IntuneDevices"
                if(!($APIntuneDevicesID)){
                    continue
                }
        
                $MobileAppAssignObj = @{
                    "mobileAppAssignments" = @(
                        @{
                            "@odata.type" = "#microsoft.graph.mobileAppAssignment"
                            "target" = @{
                                "@odata.type" = "#microsoft.graph.groupAssignmentTarget"
                                "groupId" = $APIntuneDevicesID
                            }
                            "intent" = "Required"
                            "settings" = @{
                                "@odata.type" = "#microsoft.graph.win32LobAppAssignmentSettings"
                                "notifications" = "showAll"
                                "installTimeSettings" = $null
                                "restartSettings" = $null
                                "deliveryOptimizationPriority" = "notConfigured"
                            }
                        }
                    )
                }

                #Assign the app
                Invoke-IntuneUploadConfigurationRequest -EndpointURL  "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$($AppConfigID)/assign" -JSONObj ($MobileAppAssignObj | ConvertTo-Json -Depth 5)
			}
        }
    }catch{
        Add-ToLog -Message $_.Exception.Message -GeneratedErrorMessage
    } 
}
###############################################End of Upload Intune win files to Azure####################################################################

Remove-Item ".\IntuneWinUpload\*"  -Recurse -Force #Remove all the intunewin files

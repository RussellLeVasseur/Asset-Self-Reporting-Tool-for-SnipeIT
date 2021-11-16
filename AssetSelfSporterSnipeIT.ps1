#####################################################################################################################################
# Configs
#####################################################################################################################################


# This script creates directories a Directory Structure for Year and Month.
$LogFileDirectory = "";

# Creates and Stores Data in an Excel Spreadsheet
$CsvFilePath = "";

# Send Emails when Critical Data has seen an Update or when new Assets report
$EmailParams = @{
    From = "";
    To = "";
    SMTPServer = "";
    port = "25";
}

# Dell Warranty API Fetch
$DellApi =@{
    Key = "";
    Secret = "";
}

# If you want to use this to update SnipeIT, you will need to make script changes to the custom fields in the code below.
$Snipe = @{
    Url = "";
    Token = "";
    DefStatusID = 8;
    WorkstationCatID = 12;
    ServerCatID = 8;
    FieldSetID = 3;
}

# Array of Software that should be checked for license/version
$SoftwareChecklist = @(
    'Stata'
    'SAS 9'
    'SPSS'
    'HLM'
    'EndNote'
    'ArcGIS'
    'Stat/Transfer'
    'Papers'
    'Nvivo'
    'MPlus'
);

# Array of Room Number that should power on daily.
$DailyPowerOnList = @(

);


#####################################################################################################################################
# Static Variables 
#####################################################################################################################################
Clear-Host;

Remove-Variable -Name DataHashTable -ErrorAction 'SilentlyContinue';
Remove-Variable -Name DataObject -ErrorAction 'SilentlyContinue';
Remove-Variable -Name Record -ErrorAction 'SilentlyContinue';
Remove-Variable -Name StataSerial  -ErrorAction 'SilentlyContinue';

$Today = Get-Date -UFormat "%d-%b-%Y";
$Year = Get-Date -Format yyyy;
$Month = Get-Date -UFormat "%m-%B";
$LogFileDate = Get-Date -UFormat "%d-%b-%Y";

$DeviceName = hostname;

[HashTable]$DataHashTable = @{};
$Win32_BIOS = Get-WMIObject -Class Win32_BIOS;

$DataHashTable.Add('SerialNumber', $Win32_BIOS.SerialNumber);
$Win32_ComputerSystem = Get-WmiObject -Class Win32_ComputerSystem;
$CsvFile = "$CsvFilePath\$($Win32_BIOS.SerialNumber).csv";

$LogFileDirectory = "$LogFileDirectory\$Year\$Month\$($Win32_BIOS.SerialNumber)-$($DeviceName)";
$LogFile = "$LogFileDirectory\$($Win32_BIOS.SerialNumber)_$($DeviceName)_$($LogFileDate)_SelfReport.log";

$StringHasher = [System.Security.Cryptography.HashAlgorithm]::Create('sha256');

Remove-Item C:\tech -Recurse -Force -ErrorAction SilentlyContinue;

$EmailParams.Add('Subject', '');
$EmailParams.Add('Body', '');

$CustomValues = @{};


###############################################################################################################################################################################################
# Functions
#####################################################################################################################################


If (!(Test-Path -Path "$($LogFileDirectory)")) { New-Item -ItemType Directory -Path "$($LogFileDirectory)"; }
If (!(Test-Path -Path $LogFile -PathType Leaf)) {
    New-Item -ItemType "file" -Path "$LogFile" -Force;
    Add-Content $LogFile "[$Date] Log File Created.";
}

Function WriteLog {
	param( [String] $Log, [Object[]] $Data )
    $Date = ((Get-Date -UFormat "%d-%b-%Y_%T") -replace ':', '-');
    Switch -WildCard ($Log) {
        "*success*" { Write-Host "[$Date] $Log" -f "Green"; }
        "*ERROR*" { Write-Host "[$Date] $Log" -f "Red"; }
        "*NEW*" { Write-Host "[$Date] $Log" -f "Yellow"; }
        Default { Write-Host "[$Date] $Log" -f "Magenta"; }
    }
    If ($Data) { $Data = ($Data | Out-String).Trim().Split("`n") | ForEach-Object { Write-Host "`t$_"}; }
    If ($Log) { Add-Content $LogFile "[$Date] $Log"; }
    If ($Data) {
        ($Data | Out-String).Trim().Split("`n") | ForEach-Object { Add-Content $LogFile ("`t" + "$_".Trim()) };
    }
}

Function EmailAlert {
    param( [String] $Subject, [String] $Body )
    $EmailParams.Subject = "$($Subject) on $($DeviceName)";
    $EmailParams.Body = "Device: $($DeviceName)`n`n$($Body)";
    Send-MailMessage @EmailParams;
}
Function GetHRSize {
    param(
        [Parameter(Mandatory=$True, ValueFromPipeline=$True)]
        [INT64] $bytes
    )
    Process {
        If ( $bytes -gt 1pb ) { "{0:N1}PB" -f ($bytes / 1pb) }
        ElseIf ( $bytes -gt 1tb ) { "{0:N1}TB" -f ($bytes / 1tb) }
        ElseIf ( $bytes -gt 1gb ) { "{0:N1}GB" -f ($bytes / 1gb) }
        ElseIf ( $bytes -gt 1mb ) { "{0:N1}MB" -f ($bytes / 1mb) }
        ElseIf ( $bytes -gt 1kb ) { "{0:N1}KB" -f ($bytes / 1kb) }
        Else   { "{0:N} Bytes" -f $bytes }
    }
}


#####################################################################################################################################
# Requirements 
#####################################################################################################################################

If (-NOT (Get-PackageProvider -ListAvailable -Name NuGet -ErrorAction SilentlyContinue)) {
    Write-Host "NuGet Package not found. Installing...";
    Install-PackageProvider NuGet -Confirm:$false -Force:$true;
}

If ($Win32_ComputerSystem.Model -eq "Virtual Machine") {
    $RequiredModules = "ImportExcel", "SnipeitPS";
} Else { $RequiredModules = "ImportExcel", "SnipeitPS", "DellBIOSProvider", "ActiveDirectory"; }

$RequiredModules | ForEach-Object {
    Try {
        $Mdle = $_;
        WriteLog -Log "Checking for $Mdle...";
        If (!(Get-Module -ListAvailable -Name $Mdle)) {
            WriteLog -Log "$Mdle not found. Installing...";
            If ($_ -eq 'ActiveDirectory') {
                Add-WindowsCapability –Online –Name “Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0”;
                Install-WindowsFeature RSAT-AD-PowerShell;
            } Else { Install-Module -Name $Mdle -Force; }
        } Else {
            $Latest = [String](Find-Module -Name SnipeitPS | Sort-Object Version -Descending)[0].version;
            $Installed = [String](Get-Module -ListAvailable SnipeitPS | Select-Object -First 1).version;
            If ([System.Version]$Latest -gt [System.Version]$Installed) {
                WriteLog -Log "[UPDATE] Updating $($Mdle)..." -Data $_;
                Update-Module SnipeitPS -Force;
            }
            Try { Import-Module -Name $Mdle -Force; }
            Catch {
                WriteLog -Log "[ERROR] Unable to Import $($Mdle) Module." -Data $_;
                EmailAlert -Subject "[ERROR] Importing Module" -Body $_;
            }
        }
    } Catch { WriteLog -Log "[ERROR] $($_ | Out-String)"; }
}
WriteLog -Log "Requirements Installed and Loaded.";

Connect-SnipeitPS -URL $Snipe.Url -apiKey $Snipe.Token;

#####################################################################################################################################
# General Device Information
#####################################################################################################################################
WriteLog -Log "Gathering Device Information...";

$Location = "$(($DeviceName).Split("-")[0])-$(($DeviceName).Split("-")[1])";

$DataHashTable.Add('Location', $Location);
$DataHashTable.Add('DeviceName', $($DeviceName));
$DataHashTable.Add('LastReported', (Get-Date));
$DataHashTable.Add('LastReportedUnix', ([Math]::Round((Get-Date -UFormat %s),0)));
$DataHashTable.Add('Model', $Win32_ComputerSystem.Model);
$DataHashTable.Add('Manufacturer', "$($Win32_ComputerSystem.Manufacturer -replace " Inc.", '')");

$DataHashTable.Add('Bios', $Win32_BIOS.SMBIOSBIOSVersion);


#####################################################################################################################################
# Operating System Information
#####################################################################################################################################
WriteLog -Log "Gathering Operating System Information...";

$Win32_OperatingSystem = Get-WmiObject -Class Win32_OperatingSystem;
$DataHashTable.Add('OS', ($Win32_OperatingSystem.Name).Split("|")[0]);
$DataHashTable.Add('Build', $Win32_OperatingSystem.Version);
If ($DataHashTable['OS'] -Contains "Server") { $ModelCatID = $Snipe.ServerCatID; }


#####################################################################################################################################
# Bios Information
#################################f###################################################################################################
WriteLog -Log "Gathering Bios Information...";

If (-NOT ($Win32_BIOS.SerialNumber)) { EmailAlert -Subject "No BIOS Serial Number" -Body ($Win32_BIOS | Out-String); }
Try {
    If ($DataHashTable['Manufacturer'] -eq 'Dell' -AND (Get-Item -Path "DellSmbios:\" -ErrorAction SilentlyContinue)) {
        Function Set-DellBiosSetting {
            param( [Object[]] $Setting, [String] $Value )
            $CurrentValue = (Get-Item -Path "DellSmbios:\$($Setting.PSChildName)\$($Setting.Attribute)").CurrentValue;
            If ($CurrentValue -ne $Value) {
                Try { 
                    Set-Item -Path "DellSmbios:\$($Setting.PSChildName)\$($Setting.Attribute)" -Value $Value;
                    WriteLog -Log "Set Bios Setting: $($Setting.Attribute) to $($Value)."
                } Catch {  WriteLog -Log "Failed to Set Bios Setting." -Data $_; }
            }    
        }

        ForEach ($Category in (Get-ChildItem -Path "DellSmbios:\").Category) {
            $CategorySettings = Get-ChildItem -Path "DellSmbios:\$($Category)" -WarningAction SilentlyContinue | Select-Object Attribute,CurrentValue,PSChildName;
            ForEach ($Setting in $CategorySettings) {
                If ($DataHashTable['BootPathSecurity'] -eq 'UEFI') {
                    Switch ($Setting.Attribute) {
                        "BootList" { $DataHashTable.Add('BootMode', (Get-Item -Path "DellSmbios:\$($Setting.PSChildName)\$($Setting.Attribute)").CurrentValue); }
                        "LegacyOrom" { $DataHashTable.Add('LegacyRoms', (Get-Item -Path "DellSmbios:\$($Setting.PSChildName)\$($Setting.Attribute)").CurrentValue); }
                        "AttemptLegacyBoot" { $DataHashTable.Add('LegacyBoot', (Get-Item -Path "DellSmbios:\$($Setting.PSChildName)\$($Setting.Attribute)").CurrentValue); }
                        "SecureBoot" { $DataHashTable.Add('SecureBoot', (Get-Item -Path "DellSmbios:\$($Setting.PSChildName)\$($Setting.Attribute)").CurrentValue); }
                    }
                }
                Switch ($Setting.Attribute) {
                    "MemorySpeed" { $MemorySpeed = $Setting.CurrentValue; }
                    "MemoryTechnology" { $MemoryType = $Setting.CurrentValue; }
                    "BootList" { $DataHashTable.Add('BootMode', (Get-Item -Path "DellSmbios:\$($Setting.PSChildName)\$($Setting.Attribute)").CurrentValue); }
                    "LegacyOrom" { $DataHashTable.Add('LegacyRoms', (Get-Item -Path "DellSmbios:\$($Setting.PSChildName)\$($Setting.Attribute)").CurrentValue); }
                    "AttemptLegacyBoot" { $DataHashTable.Add('LegacyBoot', (Get-Item -Path "DellSmbios:\$($Setting.PSChildName)\$($Setting.Attribute)").CurrentValue); }
                    "SecureBoot" { $DataHashTable.Add('SecureBoot', (Get-Item -Path "DellSmbios:\$($Setting.PSChildName)\$($Setting.Attribute)").CurrentValue); }
                    "UefiBootPathSecurity" { $DataHashTable.Add('BootPathSecurity', (Get-Item -Path "DellSmbios:\$($Setting.PSChildName)\$($Setting.Attribute)").CurrentValue); }
                    "EmbNic1" { Set-DellBiosSetting -Setting $Setting -Value "Enabled"; }
                    "UefiNwStack" { Set-DellBiosSetting -Setting $Setting -Value "Enabled"; }
                    "SmartErrors" { Set-DellBiosSetting -Setting $Setting -Value "Enabled"; }
                    "SfpNic" { Set-DellBiosSetting -Setting $Setting -Value "EnabledPXE"; }
                    "TpmSecurity " { Set-DellBiosSetting -Setting $Setting -Value "Enabled"; }
                    "TpmActivation" { Set-DellBiosSetting -Setting $Setting -Value "Enabled"; }
                    "AcPwrRcvry" { Set-DellBiosSetting -Setting $Setting -Value "Last"; }
                    "DeepSleepCtrl" { Set-DellBiosSetting -Setting $Setting -Value "Disabled"; }
                    "WakeOnLan" { Set-DellBiosSetting -Setting $Setting -Value "Disabled"; }
                    "BlockSleep" { Set-DellBiosSetting -Setting $Setting -Value "Disabled"; }
                    "WirelessLan" { If (-NOT (Get-WmiObject -Class win32_battery)) { Set-DellBiosSetting -Setting $Setting -Value "Disabled"; } }
                    "BluetoothDevice" { If (-NOT (Get-WmiObject -Class win32_battery)) { Set-DellBiosSetting -Setting $Setting -Value "Disabled"; } }
                    "AutoOn" {
                        If ($DailyPowerOnList -Contains $Location) { Set-DellBiosSetting -Setting $Setting -Value "EveryDay"; }
                        Else { Set-DellBiosSetting -Setting $Setting -Value "SelectDays"; } Break;
                    }
                    "AutoOnHr" { Set-DellBiosSetting -Setting $Setting -Value "7"; Break; }
                    "AutoOnMn" { Set-DellBiosSetting -Setting $Setting -Value "0"; Break; }
                    "AutoOnTue" { If ($DailyPowerOnList -NotContains $Location) { Set-DellBiosSetting -Setting $Setting -Value "Enabled"; Break; } }
                }
            }
        }
    }
} Catch { 
    WriteLog -Log "[ERROR] Issue Configuring Bios" -Data $_;
    WriteLog -Log "$($_ | Out-String)"; 
}


#####################################################################################################################################
# Network Adapter Configurations
#####################################################################################################################################
WriteLog -Log "Gathering Network Adapter Information...";

$Win32_NetworkAdapterConfiguration = Get-WmiObject -ClassName Win32_NetworkAdapterConfiguration -Filter 'NOT Description LIKE "%Loopback%"' | Where-Object { $_.IPAddress };
$DataHashTable.Add('IpAddress', "$($Win32_NetworkAdapterConfiguration.IPAddress | Where-Object { $_ -notlike "*:*"; })");
$DataHashTable.Add('MacAddress', "$($Win32_NetworkAdapterConfiguration.MacAddress)");
Switch ($Win32_NetworkAdapterConfiguration.DHCPEnabled) {
    "True" { $DataHashTable.Add('DHCP', "Enabled"); }
    "False" { $DataHashTable.Add('DHCP', "Disabled"); }
}


#####################################################################################################################################
# Group Access
#####################################################################################################################################
WriteLog -Log "Gathering Local Group Information...";

$LocalAdministrators = Get-LocalGroupMember -Group "Administrators";
$DataHashTable.Add('LocalAdmins', ($LocalAdministrators).Name -join "`n");

$RemoteDesktopUsers = Get-LocalGroupMember -Group "Remote Desktop Users";
$DataHashTable.Add('RemoteUsers', ($RemoteDesktopUsers).Name -join "`n");


#####################################################################################################################################
# Uptime
#####################################################################################################################################
WriteLog -Log "Calculating Uptime...";

$Uptime = "";
$UptimeVal = ((Get-Date)-($Win32_OperatingSystem).ConvertToDateTime($Win32_OperatingSystem.LastBootUpTime));
Switch ($true) {
    ($UptimeVal.Days -gt 0) { $Uptime += "$($UptimeVal.Days)D:"; }
    ($UptimeVal.Hours -gt 0) { $Uptime += "$($UptimeVal.Hours)H:"; }
    ($true) { $Uptime += "$($UptimeVal.Minutes)M:$($UptimeVal.Seconds)S"; $DataHashTable.Add('Uptime', $Uptime); }
}


#####################################################################################################################################
# Software
#####################################################################################################################################
WriteLog -Log "Gathering Software Information...";

$Software = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*).DisplayName
$Software += (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*).DisplayName
$Software = $Software | Sort-Object | Select-Object -Unique;
$DataHashTable.Add('Software', ($Software -join ';'));

$SoftwareHash = $StringHasher.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($DataHashTable['Software']));
$DataHashTable.Add('SoftwareHash', ([System.BitConverter]::ToString($SoftwareHash)).Replace('-', ''));

$LicensedSoftware = @();
$SoftwareChecklist | ForEach-Object {
    $SW = $_;
    $Installed = ($Software | Where-Object { $_ -like "$($SW)*" -AND $_ -notlike "background" });
    If ($Installed.Count -gt 1) {
        EmailAlert -Subject "Multiple Licensed Software Versions Installed" -Body "$($Installed)";
    } ElseIf ($Installed) {  
        $LicensedSoftware += $Installed;
    }
}
$DataHashTable.Add('LicensedSoftware', ($LicensedSoftware -join "`n"));



#####################################################################################################################################
# Drive Configuration Collection 
#####################################################################################################################################
WriteLog -Log "Gathering Internal and External Drive Information...";

Remove-Variable -Name InternalMedia -ErrorAction SilentlyContinue;
Remove-Variable -Name RemovableMedia -ErrorAction SilentlyContinue;
  
$DiskDrives = Get-WmiObject Win32_DiskDrive -Property * | Sort-Object Size;
$DiskVolumes = Get-Volume | Sort-Object Index;
$PhysicalDisks = Get-PhysicalDisk;
$HighDriveUsage = $false;
$InternalDisks = @();
$InternalMedia = @();
$RemovableMedia = @();
ForEach ($Disk in $DiskDrives) {
    If ($Disk.HealthStatus -ne 'Healthy') {
        ###############
        ###############
        ###############
    }
    Remove-Variable -Name DiskInfo -ErrorAction SilentlyContinue;
    Remove-Variable -Name DriveType -ErrorAction SilentlyContinue;
    $PhysicalDisk = ($PhysicalDisks | Where-Object { $_.DeviceID -eq (($Disk.DeviceID).substring((($Disk.DeviceID).Length)-1)) });
    If ($PhysicalDisk.MediaType -eq 'Unspecified') {
        If ($PhysicalDisk.CannotPoolReason -eq 'Removable Media' -AND $PhysicalDisk.BusType -eq 'USB') { $Disk.MediaType = 'USB'; }
    } Else { $Disk.MediaType = $PhysicalDisk.MediaType; }
    If ($PhysicalDisk.BusType -eq $Disk.MediaType) { $DiskType = $PhysicalDisk.BusType; } 
    Else { $DiskType = "$($Disk.MediaType)-$($PhysicalDisk.BusType)"; }
    $DiskInfo = "Disk$($Disk.Index): [$($DiskType)] $($Disk.Model) ($(GetHRSize $Disk.size))`n";
    $PartitionQuery = 'ASSOCIATORS OF {Win32_DiskDrive.DeviceID="'+$($Disk.DeviceID.replace('\','\\'))+'"} WHERE AssocClass=Win32_DiskDriveToDiskPartition';
    $WmiPartitions = @(Get-WmiObject -Query $PartitionQuery | Sort-Object StartingOffset);
    ForEach ($Partition in $WmiPartitions) {
        $DiskInfo += "---- Part$($Partition.Index): $(GetHRSize $Partition.Size) $($Partition.Type)`n";
        $VolumeQuery = 'ASSOCIATORS OF {Win32_DiskPartition.DeviceID="'+$Partition.DeviceID+'"} WHERE AssocClass=Win32_LogicalDiskToPartition';
        $WmiVolumes   = @(Get-WmiObject -Query $VolumeQuery);
        ForEach ($Volume in $WmiVolumes) {
            $VolumeData = "$($Volume.name) [$($Volume.FileSystem)] $((GetHRSize ($Volume.Size - $Volume.FreeSpace)) -replace "GB",'')/$(GetHRSize $Volume.Size) ($(GetHRSize ($Volume.FreeSpace)) Free)"
            If ($Volume.name -eq 'C:') {
                $DataHashTable.Add('BootDrive', "$($Volume.name) [$($DiskType)] $((GetHRSize ($Volume.Size - $Volume.FreeSpace)) -replace "GB",'')/$(GetHRSize $Volume.Size) ($(GetHRSize ($Volume.FreeSpace)) Free)");
                $DataHashTable.Add('HasSSD', ('Yes','No')[($Disk.MediaType -ne 'SSD')]);
            }
            $DiskVolume = ($DiskVolumes | Where-Object { $_.Driveletter -eq ($Volume.DeviceID -replace ":",'') });
            If ($PhysicalDisk.BusType -eq 'USB') { $DriveType = 'Removable'; } 
            Else { $DriveType = $DiskVolume.DriveType; }
            $DiskInfo += "--------- $VolumeData`n";
            If (($DiskVolume.SizeRemaining / $DiskVolume.Size) -lt .1 -AND $DriveType -ne 'Removable') { $HighDriveUsage = $true; }
        }
    }
    Switch ($DriveType) {
        "Fixed" { 
            $InternalDisks += "[$($Disk.MediaType)] $($Disk.Model) ($(GetHRSize $Disk.size))";
            $InternalMedia += $DiskInfo.Trim(); Break; 
        }
        "Removable" { $RemovableMedia += $DiskInfo.Trim(); Break; }
    }
}
$DataHashTable.Add('Drives', $InternalDisks -join "`n");
$DataHashTable.Add('InternalMedia', $InternalMedia -join "`n");
$DataHashTable.Add('RemovableMedia', $RemovableMedia -join "`n");
If ($HighDriveUsage) { EmailAlert -Subject "Drive Usage Very High" -Body "$($InternalMedia)"; }


#####################################################################################################################################
# Logged In Users
#####################################################################################################################################
WriteLog -Log "Gathering Active User Information...";

Try {
    $ActiveUsers = @();
    $LoggedInUsers = quser | ForEach-Object -Process { $_ -replace '\s{2,}',','; };
    $LoggedInUsers = $LoggedInUsers | ConvertFrom-Csv;
    ForEach ($User in $LoggedInUsers) {
        Switch -Wildcard ($User.SESSIONNAME) {
            "console" { $Session = "Local - $($User."LOGON TIME")"; }
            "rdp*" { $Session = "RDP - $($User."LOGON TIME")"; }
            default { $Session = $User.SESSIONNAME; }
        }
        $ActiveUsers +=  "$($User.USERNAME -replace '>', '') ($($Session))";
    }
    $DataHashTable.Add('ActiveUsers', ($ActiveUsers -join "`n"));
} Catch {
   WriteLog -Log "[ERROR] Unable to Collect User Information." -Data $_;
   EmailAlert -Subject "[ERROR] Unable to Collect User Information." -Body $_;
}


#####################################################################################################################################
# Webcam IdentIfication
#####################################################################################################################################
WriteLog -Log "Gathering Webcam Information...";

$Webcam = (Get-PnpDevice | Where-Object  { $_.Class -eq "Image" -AND $_.FriendlyName -like "*webcam*" } |
            Get-PnpDeviceProperty | Where-Object { $_.KeyName -eq "DEVPKEY_Device_FriendlyName" }).Data -join "`n";
Switch ($true) {
    ($null -ne $Webcam) { $DataHashTable.Add('Webcam', "$($Webcam)"); Break; } 
    default { $DataHashTable.Add('Webcam', ''); }
}

#####################################################################################################################################
# GPU IdentIfication
#####################################################################################################################################
WriteLog -Log "Gathering Graphics Information...";

$GraphicsCard = (Get-PnpDevice | Where-Object {$_.Class -eq "Display" -AND $_.Status -eq 'OK'} | 
                Get-PnpDeviceProperty | Where-Object { $_.Keyname -eq "DEVPKEY_NAME" } | Sort-Object -Property Data).Data -join "`n";
Switch ($true) {
    ($null -ne $GraphicsCard) { $DataHashTable.Add('Graphics', "$($GraphicsCard)"); Break; } 
    default { $DataHashTable.Add('Graphics', ''); }
}


#####################################################################################################################################
# RAM/Memory
#####################################################################################################################################
WriteLog -Log "Gathering Memory Information...";

$Memory = Get-WmiObject -Class Win32_PhysicalMemory;
$MemoryVoltage = $Memory[0].ConfiguredVolage;
If (-NOT ($MemorySpeed)) { $MemorySpeed = "$($Memory[0].Speed)MHz"; }
If (-NOT ($MemoryType)) { 
    Switch ($MemoryVoltage) { 
        '1200' { $MemoryType = "DDR4"; } 
        '1500' { $MemoryType = "DDR3"; } 
        default { $MemoryType = ''; } 
    }
}
$Memory = $Memory | Measure-Object -Property Capacity -Sum;

$MemoryAvailable = [math]::Round(($Win32_OperatingSystem.FreePhysicalMemory / 1MB),1);
$MemoryUsed = [math]::Round((($Memory.Sum / 1gb)-($Win32_OperatingSystem.FreePhysicalMemory / 1MB)),1);
$MemoryInstalled = "$($Memory.Sum / 1gb)GB";

$DataHashTable.Add('RAM', "$($MemoryUsed)/$($MemoryInstalled) [$($Memory.Count)] $($MemorySpeed) $($MemoryType)");
$DataHashTable.Add('RAM_Installed', "$($MemoryInstalled) [$($Memory.Count)]");

If ([int]$MemoryAvailable -lt 1 -AND $DataHashTable['Model'] -ne "Virtual Machine") {
    EmailAlert -Subject "Low RAM Availability" -Body "Current Available RAM: $($MemoryAvailable)Gb";
}


#####################################################################################################################################
# CPU/Processor
#####################################################################################################################################
WriteLog -Log "Gathering Processor Information...";

$Win32_Processor = (Get-WmiObject Win32_Processor | Select-Object *);
If ($Win32_Processor.Count -gt 1) { 
    $Win32_Processor = $Win32_Processor[0]; 
    $Win32_Processor.Name = "[2] $($Win32_Processor.Name)"; 
}
$AssetProcessor = ($Win32_Processor.Name -replace '\(TM\)|\(R\)','');
$AssetProcessor = ($AssetProcessor -replace '@',"$($Win32_Processor.NumberOfCores)c/$($Win32_Processor.NumberOfLogicalProcessors)t");
$AssetProcessor = ($AssetProcessor -replace '  | 0 ',' ');
$DataHashTable.Add('CPU', $AssetProcessor);


#####################################################################################################################################
# Load Record If Exists
#####################################################################################################################################
WriteLog -Log "Loading Existing Record...";

If (Test-Path -Path $CsvFile -PathType Leaf) {
    $Record = Import-Csv -Path $CsvFile;
}


#####################################################################################################################################
# Fetch Dell Warranty Data
#####################################################################################################################################
WriteLog -Log "Checking Dell Warranty Information...";

If ($DataHashTable['Model'] -ne "Virtual Machine") {
    If (!$Record -OR !$Record.Purchased -OR !$Record.WarrantyMonths) {
        WriteLog -Log "Updating Purchase and Warranty Dates..."
        If ($DataHashTable['Manufacturer'] -eq "Dell") {
            WriteLog -Log "[Dell Warranty] No Dell Information in the Record. Fetching Purchase Date and Warranty Date.";
            Try {
                $AuthURI = "https://apigtwb2c.us.dell.com/auth/oauth/v2/token";
                $OAuth = "$($DellApi.Key)`:$($DellApi.Secret)";
                $Bytes = [System.Text.Encoding]::ASCII.GetBytes($OAuth);
                $EncodedOAuth = [Convert]::ToBase64String($Bytes);
                $DellTokenHeaders = @{ };
                $DellTokenHeaders.Add("authorization", "Basic $EncodedOAuth");
                $Authbody = 'grant_type=client_credentials';
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;
                $AuthResult = Invoke-RESTMethod -Method Post -Uri $AuthURI -Body $AuthBody -Headers $DellTokenHeaders;
                $Token = $AuthResult.access_token;
                $DellApiHeaders = @{ 
                    "Accept" = "application/json"
                    "Authorization" = "Bearer $($Token)"
                };
                $DellApiParams = @{ };
                $DellApiParams = @{servicetags = $DataHashTable['SerialNumber']; Method = "GET"};
                $DellApiResponse = Invoke-RestMethod -Uri "https://apigtwb2c.us.dell.com/PROD/sbil/eapi/v5/asset-entitlements" -Headers $DellApiHeaders -Body $DellApiParams -Method Get -ContentType "application/json" -ea 0;
                $DellApiResponse = $DellApiResponse | ConvertTo-Json | ConvertFrom-Json;
                If ($DellApiResponse) {
                    $DataHashTable.Add('WarrantyExpiration', (($DellApiResponse.entitlements | Select-Object -Last 1).endDate | Get-Date));
                    $DataHashTable.Add('Purchased', ($DellApiResponse.shipDate | Get-Date -UFormat "%Y-%m-%d"));
                    $DataHashTable.Add('Age', [math]::Round(((New-TimeSpan -Start $DataHashTable['Purchased'] -End $Today).Days / 365), 1));
                    If ($DellApiResponse.ProductID -Like '*desktop*') { 
                        $DellApiResponse.ProductID = 'Desktop';
                    } ElseIf ($DellApiResponse.ProductID -Like '*laptop*') { 
                        $DellApiResponse.ProductID = 'Laptop';
                    } ElseIf ($DellApiResponse.ProductID -Like '*server*') { 
                        $DellApiResponse.ProductID = 'Server';
                    }
                    If ((-NOT ($DataHashTable['Purchased'])) -OR (-NOT ($DataHashTable['WarrantyExpiration']))) {
                        WriteLog -Log "[ERROR] Potential problem with Dell Warranty Fetch." -Data $DellApiResponse;
            }}} Catch {
                WriteLog -Log "[ERROR] Error Obtaining Dell Warranty Information." -Data $_;
            } 
        }
    } Else {
        $DataHashTable.Add('Purchased', ($Record.Purchased | Get-Date -UFormat "%Y-%m-%d"));
        $DataHashTable.Add('WarrantyExpiration', $Record.WarrantyExpiration);
        $DataHashTable.Add('Age', [math]::Round(((New-TimeSpan -Start $DataHashTable['Purchased'] -End $Today).Days / 365), 1));
    }
    $DataHashTable.Add('WarrantyMonths', [math]::Round(((New-TimeSpan -Start $DataHashTable['Purchased'] -End $DataHashTable['WarrantyExpiration']).Days) / 30.33));
}

#####################################################################################################################################
# Update SnipeIT 
#####################################################################################################################################
WriteLog -Log "Checking in to SnipeIT...";
$SnipeAsset = Get-SnipeItAsset -asset_serial $DataHashTable['SerialNumber'];
$CustomValues.Add('purchase_date', $DataHashTable['Purchased']);
$CustomValues.Add('warranty_months', $DataHashTable['WarrantyMonths']);
$CustomValues.Add('_snipeit_mac_address_1', $DataHashTable['MacAddress']);
$CustomValues.Add('_snipeit_cpu_2', $DataHashTable['CPU']);
$CustomValues.Add('_snipeit_ram_3', $DataHashTable['RAM']);
$CustomValues.Add('_snipeit_fqdn_4', [System.Net.Dns]::GetHostEntry($DataHashTable['IpAddress']).HostName);
$CustomValues.Add('_snipeit_operating_system_5', $DataHashTable['OS']);
$CustomValues.Add('_snipeit_ip_address_9', $DataHashTable['IpAddress']);
$CustomValues.Add('_snipeit_bios_11', $DataHashTable['Bios']);
$CustomValues.Add('_snipeit_last_reported_12', (Get-Date -UFormat "%d-%b-%Y %T"));
$CustomValues.Add('_snipeit_graphics_13', $DataHashTable['Graphics']);
$CustomValues.Add('_snipeit_boot_drive_15', $DataHashTable['BootDrive']);
$CustomValues.Add('_snipeit_internal_media_16', $DataHashTable['InternalMedia']);
$CustomValues.Add('_snipeit_external_media_17', $DataHashTable['RemovableMedia']);
$CustomValues.Add('_snipeit_licensed_software_18', $DataHashTable['LicensedSoftware']);
$CustomValues.Add('_snipeit_remote_desktop_users_19', $DataHashTable['RemoteUsers']);
If (!$SnipeAsset) {
    Try {
        Try {
            $Manufacturer = Get-SnipeItManufacturer -search $DataHashTable['Manufacturer'];
            If (!$Manufacturer) { $Manufacturer = New-SnipeItManufacturer -name $DataHashTable['Manufacturer']; }
            $ManufacturerID = $Manufacturer.id;
        } Catch { WriteLog -Log "[SnipeIT] [ERROR] Unable to obtain Manufacturer ID." -Data $_; }
        Try {
            $Model = Get-SnipeItModel -all | Where-Object { $_.name -eq "$($DataHashTable['Model'])" };
            $ModelData = $Model.notes -replace "&quot;",'"' | ConvertFrom-Json;
            If ($ModelData.LatestBios -gt $DataHashTable['Bios']) {
                ##########################################
            }
            If ($Model.total -eq 0) {
                If ($DataHashTable['OS'] -Contains "Server") { $ModelCatID = $Snipe.ServerCatID; } 
                Else { $ModelCatID = $Snipe.WorkstationCatID; }
                $Model = New-SnipeItModel -name $DataHashTable['Model'] -manufacturer_id $ManufacturerID -fieldset_id $Snipe.FieldSetID -category_id $ModelCatID;
            }
        } Catch { WriteLog -Log "[SnipeIT] [ERROR] Unable to obtain Model ID." -Data $_; }
        $SnipeAsset = New-SnipeItAsset -name $DataHashTable['DeviceName'] -status_id 8 -model_id $Model.id -serial $DataHashTable['SerialNumber'] -asset_tag $DataHashTable['SerialNumber'] -customfields $CustomValues;
        WriteLog -Log "[SnipeIT] Created a new Asset in SnipeIT.";
    } Catch { WriteLog -Log "[SnipeIT] [ERROR] Unable to Create new Asset." -Data $_; }
} ElseIf ($SnipeAsset.Count -gt 1) {
    WriteLog -Log "[ERROR] Multiple Assets with Identical Serial Numbers Found in SnipeIT.";
    EmailAlert -Subject "Multiple Assets with Identical Serial Numbers Found." -Body "Asset Name: $($DeviceName)`n`n$($SnipeAsset | Format-List | Out-String)";
} Else {
    $UserAssigned = $SnipeAsset.assigned_to.username;
    $RemoteUsers = ($RemoteDesktopUsers | Where-Object { $_.ObjectClass -ne 'Group' } | Select-Object -Property @{ Name="Name"; Expression={ ($_.Name).Split("\")[1] } }).Name;
    If ($RemoteUsers -NotContains $UserAssigned) {
        WriteLog -Log "Adding $($UserAssigned) to local Remote Desktop Users Group...";
        Add-LocalGroupMember -Group "Remote Desktop Users" -Member $UserAssigned;
        $DataHashTable['RemoteUsers'] = (Get-LocalGroupMember -Group "Remote Desktop Users").Name -join "`n";
    }
    If ($RemoteUsers.Count -gt 1) {
        WriteLog -Log "Found more than one user assigned to the local Remote Desktop Users Group...";
        $UnauthorizedRemoteUsers = $RemoteUsers | Where-Object { $_ -ne $UserAssigned }
        $UnauthorizedRemoteUsers | Get-ADUser | ForEach-Object {
            WriteLog -Log "Removed $($_.SamAccountName) from the local Remote Desktop Users Group...";
            Remove-LocalGroupMember -Group "Remote Desktop Users" -Member $_.SamAccountName;
            $DataHashTable['RemoteUsers'] = (Get-LocalGroupMember -Group "Remote Desktop Users").Name -join "`n";
        }
        EmailAlert -Subject "Unauthorized Remote Users" -Body "The following accounts were removed from the local Remote Desktop Users group:`n$($UnauthorizedRemoteUsers)";
    }
    If ($SnipeAsset.asset_tag) { $DataHashTable.Add('AssetTag', $SnipeAsset.asset_tag); }
    If ($DataHashTable['Location'] -notlike "CCCJ-*") {
        If ($SnipeAsset.location -AND $SnipeAsset.location.id) {
            $SnipeLocation = (Get-SnipeItLocation -id $SnipeAsset.location.id).name;
        } Else { $SnipeLocation = "UNASSIGNED"; }
        $Locations = Get-SnipeItLocation -search $DataHashTable['Location'];
        ForEach ($L in $Locations) { If ($L.name -eq $DataHashTable['Location']) { $Location = $L; } }
        If ($SnipeLocation -ne "UNASSIGNED" -AND !$Location.name) {
            WriteLog -Log "[ERROR] No Location found for Asset $($DeviceName)";
            EmailAlert -Subject "New Location Created in SnipeIT." -Body "Asset Name: $($DeviceName)`nLocation: $($DataHashTable['Location'])";
            $Location = New-SnipeItLocation -name $DataHashTable['Location'];
        }
        If ((($SnipeLocation.Split('-') | Select-Object -First 2) -join '-') -ne $Location.name) {
            EmailAlert -Subject "Asset Location in SnipeIT does not match naming convention." -Body "SnipeIT Location: $($SnipeLocation)`nAsset Name: $($DeviceName)`n$($SnipeAsset | Format-List | Out-String)";
        }
    }
    $LicensedSoftware | ForEach-Object {
        $SW = $_;
        Switch -Wildcard ($SW) {
            'Stata*' {
                If (Get-ChildItem -Path "C:\Program Files\$($SW -replace ' ','')\stata.lic" -ErrorAction SilentlyContinue) {
                    $StataLicense = (Get-ChildItem -Path "C:\Program Files\$($SW -replace ' ','')\stata.lic" | Get-Content).Split('!');
                } ElseIf (Get-ChildItem -Path "C:\Program Files (x86)\$($SW -replace ' ','')\STATA.LIC" -ErrorAction SilentlyContinue) {
                    $StataLicense = (Get-ChildItem -Path "C:\Program Files (x86)\$($SW -replace ' ','')\STATA.LIC" | Get-Content).Split('!');
                }
                If (!$StataLicense) {
                    EmailAlert -Subject "Unlicensed Stata Install" -Body "No Stata license is configured for install on this PC.";
                } Else {
                    $LocalStataSerial = $StataLicense[0];
                    $AssignedLicense = Get-SnipeItLicense -asset_id $SnipeAsset.id | Where-Object { $_.name -like 'Stata*' };
                    If (!$AssignedLicense) {
                        $SnipeItStataLicense = Get-SnipeItLicense | Where-Object { $_.name -like "Stata*" -AND ([BigInt]($_.product_key.Split("`n")[0]) -eq [BigInt]$LocalStataSerial) };
                        If ($SnipeItStataLicense) {
                            $OpenStataSeats = Get-SnipeItLicenseSeat -id $SnipeItStataLicense.id | 
                                                    Where-Object { !$_.assigned_asset } | 
                                                    Sort-Object -Property id;
                            If (!$OpenStataSeats -OR $OpenStataSeats.Count -lt 1) {
                                EmailAlert -Subject "Stata License Error" -Body "The Stata license assigned has no open seats.`n`nInstalled:`n$($StataLicense | Out-String)";
                            } Else {
                                Set-SnipeItLicenseSeat -id $SnipeItStataLicense.id -seat_id $OpenStataSeats[0].id -asset_id $SnipeAsset.id;
                            }
                        } Else {
                            EmailAlert -Subject "Unknown Stata License" -Body "A Stata license was seen that does not match anything in the inventory.`n`nInstalled:`n$($StataLicense)";
                        }
                    } Else {
                        $SnipeItStataSerial = $AssignedLicense.product_key.Split("`n")[0];
                        If ([BigInt]$SnipeItStataSerial -ne [BigInt]$LocalStataSerial) {
                            EmailAlert -Subject "Mismatched Stata License" -Body "A Stata license assigned is not the same license that is instaled.`n`nInstalled:`n$($StataLicense)`n`nAssigned:`n$($AssignedLicense | Out-String)";
                        }
                    }
                }
            }
            'Stat/Transfer*' {
                $AssignedLicense = Get-SnipeItLicense -asset_id $SnipeAsset.id | Where-Object { $_.name -like 'Stat/Transfer*' };
                If (!$AssignedLicense) {
                    
                }
            }
            'HLM*' {

            }
            'MPlus*' {
                $AssignedLicense = Get-SnipeItLicense -asset_id $SnipeAsset.id | Where-Object { $_.name -like 'MPlus*' };
            }
            'SAS*' {
                $AssignedLicense = Get-SnipeItLicense -asset_id $SnipeAsset.id | Where-Object { $_.name -like 'SAS*' };
                If (!$AssignedLicense) {
                    $SnipeItSasLicense = Get-SnipeItLicense | Where-Object { $_.name -like "SAS*" };
                    If ($SnipeItSasLicense) {
                        $OpenSasSeats = Get-SnipeItLicenseSeat -id $SnipeItSasLicense.id | Where-Object { !$_.assigned_asset } | Sort-Object -Property id;
                        If (!$OpenSasSeats -OR $OpenSasSeats.Count -lt 1) {
                            EmailAlert -Subject "SAS License Error" -Body "The SAS license assigned has no open seats.";
                        } Else {
                            $AssignLicense = Set-SnipeItLicenseSeat -id $SnipeItSasLicense.id -seat_id $OpenSasSeats[0].id -asset_id $SnipeAsset.id;
                        }
                    } Else {
                        EmailAlert -Subject "Unknown SAS License" -Body "A SAS license cannot be found in the inventory.";
                    }
                }
            }
            'SPSS*' {
                $AssignedLicense = Get-SnipeItLicense -asset_id $SnipeAsset.id | Where-Object { $_.name -like 'SPSS*' };
                If (!$AssignedLicense) {
                    $SnipeItSpssLicense = Get-SnipeItLicense | Where-Object { $_.name -like "SPSS*" };
                    If ($SnipeItSpssLicense) {
                        $OpenSpssSeats = Get-SnipeItLicenseSeat -id $SnipeItSpssLicense.id | Where-Object { !$_.assigned_asset } | Sort-Object -Property id;
                        If (!$OpenSpssSeats -OR $OpenSpssSeats.Count -lt 1) {
                            EmailAlert -Subject "SPSS License Error" -Body "The SPSS license assigned has no open seats.";
                        } Else { $AssignedLicense = Set-SnipeItLicenseSeat -id $SnipeItSpssLicense.id -seat_id $OpenSpssSeats[0].id -asset_id $SnipeAsset.id; }
                    } Else { EmailAlert -Subject "Unknown SPSS License" -Body "An SPSS license cannot be found in the inventory."; }
                }
            }
            'EndNote*' {
                $AssignedLicense = Get-SnipeItLicense -asset_id $SnipeAsset.id | Where-Object { $_.name -like 'EndNote*' };
                If (!$AssignedLicense) {

                }
            }
            'ArcGIS*' {
                $AssignedLicense = Get-SnipeItLicense -asset_id $SnipeAsset.id | Where-Object { $_.name -like 'ArcGIS*' };
                If (!$AssignedLicense) {
                    
                }
            }
            'Papers*' {
                $AssignedLicense = Get-SnipeItLicense -asset_id $SnipeAsset.id | Where-Object { $_.name -like 'Papers*' };
                If (!$AssignedLicense) {

                }
            }
            'Nvivo*' {
                $AssignedLicense = Get-SnipeItLicense -asset_id $SnipeAsset.id | Where-Object { $_.name -like 'Nvivo*' };
                If (!$AssignedLicense) {

                }
            }
        }
    }

    Try {
        $Model = Get-SnipeItModel -search $DataHashTable['Model'];
        $ManufacturerID = $SnipeAsset.manufacturer.id;
        $UpdatedAsset = Set-SnipeItAsset -name $DataHashTable['DeviceName'] -id $SnipeAsset.id -status_id $Snipe.DefStatusID -customfields $CustomValues;
        WriteLog -Log "[SnipeIT] Updated an Asset in SnipeIT." -Data $UpdatedAsset;
        $SnipeAsset = $UpdatedAsset;
    } Catch { WriteLog -Log "[ERROR] Unable to Update SnipeIT Asset." -Data $_; }
}

# Check for Duplicate objects in SnipeIT
$DuplicateNames = Get-SnipeItAsset -Search $DataHashTable['DeviceName'] | Where-Object { $_.serial -ne $DataHashTable['SerialNumber'] -AND $_.assigned_to.name -notlike "$($DeviceName)*" };
If ($DuplicateNames.Count -gt 0) {
    $RemoveValues = {
        _snipeit_fqdn_4 = '';
        _snipeit_ip_address_9 = '';
    }
    ForEach ($Duplicate in $DuplicateNames) {
        #$UpdatedAsset = Set-SnipeItAsset -name '' -id $Duplicate.id -status_id 5 -customfields $RemoveValues;
    }
    WriteLog -Log "[ERROR] Duplicate Information Found in SnipeIT.";
    EmailAlert -Subject "Multiple Assets with Identical Data Found. Removed data from inventory object." -Body "Asset Name: $($DeviceName)`n`n$($DuplicateNames | Format-List | Out-String)";
}

#####################################################################################################################################
# Save New Record If Not Exists
##########################################################################
$DataObject = [PSCustomObject]$DataHashTable;
$DataObject = $DataObject | Select-Object -Property AssetTag,SerialNumber,DeviceName,Model,Age,LastReported,MacAddress,IpAddress,HasSSD,LocalAdmins,RemoteUsers,ActiveUsers,Webcam,Graphics,SecureBoot,BootDrive,InternalMedia,RemovableMedia,CPU,RAM,SoftwareHash,Software,Location,WarrantyMonths,Purchased,DHCP,OS,BootPathSecurity,LegacyBoot,LegacyRoms,Uptime,LastReportedUnix,Build,Drives,WarrantyExpiration,Bios,Manufacturer,LicensedSoftware,RAM_Installed,BootMode;

If (!$Record) {
    Try {
        $DataObject | Export-Csv -Path $CsvFile;
        WriteLog -Log "[REPORT] [NEW] $($DataObject.DeviceName) Self-Reported" -Data $DataObject;
        EmailAlert -Subject "New Asset Self-Reported" -Body ($DataObject | Format-List | Out-String);
    } Catch { WriteLog -Log "[ERROR] Error Saving Excel File.";  WriteLog -Log "[ERROR] $($_ | Out-String)"; }
} Else {

#####################################################################################################################################
# Check for Added/Removed Software 
#####################################################################################################################################
    WriteLog -Log "Checking Software...";

    If ($Record.SoftwareHash -ne $DataObject.SoftwareHash) {
        $OldSoftware = ($Record.Software).Split(";");
        $NewSoftware = ($DataObject.Software).Split(";");
        $SoftwareChange = 0;
        $SoftwareChanges = @();
        $EmailText = "";
        ForEach ($NSW in $NewSoftware) {
            If (($NSW -ne '') -AND ($OldSoftware -NotContains $NSW)) {
                $SoftwareChanges += New-Object -TypeName PSObject -Property @{ Added=$NSW; };
            }
        }
        ForEach ($OSW in $OldSoftware) {
            If (($OSW -ne '') -AND ($NewSoftware -NotContains $OSW)) {
                $SoftwareChanges += New-Object -TypeName PSObject -Property @{ Removed=$OSW; };
            }
        }
        If ($SoftwareChanges.Added.Length -gt 0) {
            $SoftwareChange = $SoftwareChange + 1;
            $EmailText += "$($SoftwareChanges | Where-Object  { $_.Added -ne $null; } | Select-Object Added | Format-Table -AutoSize | Out-String)";
        }
        If ($SoftwareChanges.Removed.Length -gt 0) {
            $SoftwareChange = $SoftwareChange + 2;
            $EmailText += "$($SoftwareChanges | Where-Object  { $_.Removed -ne $null; } | Select-Object Removed | Format-Table -AutoSize | Out-String)";
        }
        Switch ($SoftwareChange) {
            1 { $SoftwareChange = "Added"; }
            2 { $SoftwareChange = "Removed"; }
            3 { $SoftwareChange = "Added & Removed"; }
        }
        If ($SoftwareChange -gt 0) {
            WriteLog -Log "[SOFTWARE] Software Change Found!";
            EmailAlert -Subject "Software $($SoftwareChange)" -Body $EmailText;
        }
    } Else { WriteLog -Log "[SOFTWARE] No Change in Software Found."; }


#####################################################################################################################################
# Check for Major Changes Since Last Report 
#####################################################################################################################################

    Try {
        $DataComparison = Compare-Object -ReferenceObject $Record -DifferenceObject $DataObject -Property IpAddress,MacAddress,DeviceName,CPU,RAM_Installed,Drives,DHCP,OS,Bios,LocalAdmins,RemoteUsers,Graphics,Webcam -IncludeEqual
        $MajorChange = $DataComparison.SideIndicator;
        If ($MajorChange -ne "==") {
            $MajorChanges = "";
            Switch ($true) {
                ($Record.IpAddress -ne $DataObject.IpAddress) { $MajorChanges += "IP Address, "; }
                ($Record.MacAddress -ne $DataObject.MacAddress) { $MajorChanges += "MAC Address, "; }
                ($Record.DeviceName -ne $DataObject.DeviceName) { $MajorChanges += "Device Name, "; }
                ($Record.CPU -ne $DataObject.CPU) { $MajorChanges += "CPU, "; }
                ($Record.RAM_Installed -ne $DataObject.RAM_Installed) { $MajorChanges += "RAM, "; }
                ($Record.Drives -ne $DataObject.Drives) { $MajorChanges += "Internal Drives, "; }
                ($Record.DHCP -ne $DataObject.DHCP) { $MajorChanges += "DHCP, "; }
                ($Record.OS -ne $DataObject.OS) { $MajorChanges += "OS, "; }
                ($Record.Bios -ne $DataObject.Bios) { $MajorChanges += "Bios, "; }
                ($Record.LocalAdmins -ne $DataObject.LocalAdmins) { $MajorChanges += "Local Admins, "; }
                ($Record.RemoteUsers -ne $DataObject.RemoteUsers) { $MajorChanges += "Remote Users, "; }
                ($Record.Graphics -ne $DataObject.Graphics) { $MajorChanges += "Graphics, "; }
                ($Record.Webcam -ne $DataObject.Webcam) { $MajorChanges += "Webcam, "; }
            }
            $MajorChanges = $MajorChanges.Substring(0, ($MajorChanges.Length - 2));
            If ($MajorChanges -ne "RAM" -AND $DataObject.Model -ne "Virtual Machine") {
                WriteLog -Log "[CONFIGURATION] Major Configuration Change Found!" -Data $($DataComparison | Format-List | Out-String);
                EmailAlert -Subject "$($MajorChanges) Changed" -Body "Change:  $MajorChanges`n$($DataComparison | Format-List | Out-String)";
            }
        } Else { WriteLog -Log "[CONFIGURATION] No Major Change in Device Configuration Found.";}
    } Catch { WriteLog -Log "[ERROR] Error Comparing Configurations." -Data $_; WriteLog -Data $DataComparison; }


#####################################################################################################################################
# Save New Data to the Excel File
#####################################################################################################################################

    $Record = $DataObject;
    Try {
        $Record | Export-Csv -Path $CsvFile;
        WriteLog -Log "[REPORT] $($Record.DeviceName) Self-Reported" -Data $Record;
    } Catch { WriteLog -Log "[ERROR] Error Saving Excel File." -Data $_; }
}


#####################################################################################################################################
# WSUS Check-In
#####################################################################################################################################

Try {
    Start-Service wuauserv;
    #$UpdateSession = new-object -com "Microsoft.Update.Session";
    #$UpdateSession.CreateupdateSearcher().Search($criteria).Updates;
    Start-sleep -seconds 10;
    wuauclt /detectnow;
    (New-Object -ComObject Microsoft.Update.AutoUpdate).DetectNow();
    wuauclt /reportnow;
    c:\windows\system32\UsoClient.exe startscan;
    WriteLog -Log "[SUCCESS] WSUS Check-In was Successful.";
} Catch {  WriteLog -Log "[ERROR] WSUS Check-In was Unsuccessful." -Data $_; }


#####################################################################################################################################
# SCCM Check-In
#####################################################################################################################################

If ($DataObject.OS -notlike "*Server*") {
    Try {
        $SMSCli = [wmiclass] "root\ccm:sms_client";
        If (-NOT (Get-WmiObject -Namespace root\ccm -Class SMS_Client)) {
	        Stop-Service -Force winmgmt -ErrorAction SilentlyContinue;
   	        Set-Location  C:\Windows\System32\Wbem\;
   	        Remove-Item C:\Windows\System32\Wbem\Repository.old -Force -ErrorAction SilentlyContinue;
   	        Rename-Item Repository Repository.old -ErrorAction SilentlyContinue;
   	        Start-Service winmgmt;
        }
        If (Get-Service -Name CcmExec) {
	        $CcmExecStatus = Get-Service -Name CcmExec | ForEach-Object { $_.status; };
	        $BITSStatus = Get-Service -Name BITS | ForEach-Object { $_.status; };
	        $WuauservStatus = Get-Service -Name wuauserv | ForEach-Object { $_.status; };
	        $WinmgmtStatus = Get-Service -Name Winmgmt | ForEach-Object { $_.status; };
	        If ($CcmExecStatus -eq "Stopped") { Get-Service -Name CcmExec | Start-Service; }
	        If ($BITSStatus -eq "Stopped") { Get-Service -Name BITS | Start-Service; }
	        If ($WuauservStatus -eq "Stopped") { Get-Service -Name wuauserv | Start-Service; }
	        If ($WinmgmtStatus -eq "Stopped") { Get-Service -Name Winmgmt | Start-Service; }
	        $MachinePolicyRetrievalEvaluation = "{00000000-0000-0000-0000-000000000021}";
	        $SoftwareUpdatesScan = "{00000000-0000-0000-0000-000000000113}";
	        $SoftwareUpdatesDeployment = "{00000000-0000-0000-0000-000000000108}";
	        $MachineStatus = Invoke-WmiMethod -Namespace root\ccm -Class sms_client -Name TriggerSchedule $MachinePolicyRetrievalEvaluation;
	        $SoftwareStatus = Invoke-WmiMethod -Namespace root\ccm -Class sms_client -Name TriggerSchedule $SoftwareUpdatesScan;
	        $SoftwareDeployStatus = Invoke-WmiMethod -Namespace root\ccm -Class sms_client -Name TriggerSchedule $SoftwareUpdatesDeployment;
            Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000121}" -ErrorAction SilentlyContinue | Out-Null; 
            Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000003}" -ErrorAction SilentlyContinue | Out-Null;
            Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000010}" -ErrorAction SilentlyContinue | Out-Null;
            Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000001}" -ErrorAction SilentlyContinue | Out-Null;
            Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000022}" -ErrorAction SilentlyContinue | Out-Null;
            Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000002}" -ErrorAction SilentlyContinue | Out-Null;
            Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000031}" -ErrorAction SilentlyContinue | Out-Null;
            Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000114}" -ErrorAction SilentlyContinue | Out-Null;
            Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000111}" -ErrorAction SilentlyContinue | Out-Null;
            Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000026}" -ErrorAction SilentlyContinue | Out-Null;
            Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000027}" -ErrorAction SilentlyContinue | Out-Null;
            Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000032}" -ErrorAction SilentlyContinue | Out-Null;
	        If ($MachineStatus -AND $SoftwareStatus -AND $SoftwareDeployStatus) { 
                WriteLog -Log "[SUCCESS] SCCM Check-In Successful."; 
            } Else {
		        $SMSCli.RepairClient();
                WriteLog -Log "[ERROR] SCCM Check-In Unsuccessful.";
	        }
        } Else { WriteLog -Log "[ERROR] SCCM Does Not Appear to be Installed."; }
    } Catch { WriteLog -Log "[ERROR] Error Checking In with SCCM." -Data $_; }
}

#########################################
# Th-th-th-th-that's all folks! 
#########################################
#[Environment]::Exit(0);

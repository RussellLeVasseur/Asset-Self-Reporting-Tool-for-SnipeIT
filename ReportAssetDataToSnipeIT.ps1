param (
    [Object]$ConfigFile=''
)

Clear-Host;

########################################################################################################################################################################################################
# Remove Stale Variables 
########################################################################################################################################################################################################
Remove-Variable -Name DataHashTable -ErrorAction 'SilentlyContinue';
Remove-Variable -Name DataObject -ErrorAction 'SilentlyContinue';
Remove-Variable -Name Record -ErrorAction 'SilentlyContinue';
Remove-Variable -Name EmailParams -ErrorAction 'SilentlyContinue';
Remove-Variable -Name Config -ErrorAction 'SilentlyContinue';


########################################################################################################################################################################################################
# Static Variables 
########################################################################################################################################################################################################
If (!$ConfigFile) { Exit 1; }
$Config = (Get-Content $ConfigFile) | ConvertFrom-Json;

$EmailParams = @{
    From=$Config.EmailParams.From;
    To=$Config.EmailParams.To;
    SMTPServer=$Config.EmailParams.SMTPServer;
    port=$Config.EmailParams.Port;
}

$LocalFileDir = $Config.LocalFileDir;
$LogFileDir = $Config.LogFileDir;
$RamAlertLog = "$LocalFileDir\LowRamAlert.json";
$StorageAlertLog = "$LocalFileDir\LowStorageAlert.json";
$RecordFileDir = $Config.RecordFileDir;
$DellApi = $Config.DellApi;
$Snipe = $Config.Snipe;
$DailyPowerOnList = $Config.DailyPowerOnList;
$KeyFile = $Config.DellBios.KeyFile;
$OldPwdFile = $Config.DellBios.OldPwdFile;
$NewPwdFile = $Config.DellBios.NewPwdFile;

$StartTime = Get-Date;
$Today = Get-Date -UFormat "%d-%b-%Y";

$DeviceName = hostname;
[HashTable]$DataHashTable = @{};
$Win32_BIOS = Get-WMIObject -Class Win32_BIOS;
$SerialNumber = $Win32_BIOS.SerialNumber;
$RandomNumber = Get-Random -Minimum 0 -Maximum 300;

# List of default, erroneous, and redundant apps that may be installed that we do not need listed under "installed software".
# The script will still notify you if install status changes for these, but will not list these apps in SnipeIT.
$DefaultSoftware = @(
    "Alertus Desktop"
    "Adobe Genuine Service"
    "AMD Settings - Branding"
    "Adobe Refresh Manager"
    "ConfigMgr Client Setup Bootstrap"
    "Dropbox Update Helper"
    "Dynamic Application Loader Host Interface Service"
    "Intel(R) Chipset Device Software"
    "Intel(R) Icls"
    "Intel(R) LMS"
    "Intel(R) Management Engine Components"
    "Intel(R) Management Engine Driver"
    "Intel(R) Processor Graphics"
    "Intel(R) OEM Extension"
    "Intel(R) Rapid Storage Technology"
    "Intel(R) Serial IO"
    "Intel(R) Trusted Connect Service Client x64"
    "Intel(R) Trusted Connect Service Client x86"
    "Intel(R) Trusted Connect Services Client"
    "Intel(R) Wireless Manageability Driver"
    "Intel(R) Wireless Manageability Driver Extension"
    "Intel Optane Pinning Explorer Extensions"
    "Maxx Audio Installer (x64)"
    "Microsoft Edge"    
    "Microsoft Edge Update"    
    "Microsoft Edge WebView2 Runtime"
    "Microsoft Mouse and Keyboard Center"
    "Microsoft OneDrive"
    "Microsoft Policy Platform"
    "Microsoft Update Health Tools"
    "Microsoft VC++ redistributables repacked."
    "Microsoft Visual C++ 2010 x64 Redistributable"
    "Microsoft Visual C++ 2010 x86 Redistributable"
    "Microsoft Visual C++ 2012 Redistributable (x64)"
    "Microsoft Visual C++ 2012 Redistributable (x86)"
    "Microsoft Visual C++ 2012 x64 Additional Runtime"
    "Microsoft Visual C++ 2012 x64 Minimum Runtime"
    "Microsoft Visual C++ 2012 x86 Additional Runtime"
    "Microsoft Visual C++ 2012 x86 Minimum Runtime"
    "Microsoft Visual C++ 2013 Redistributable (x64)"
    "Microsoft Visual C++ 2013 Redistributable (x86)"
    "Microsoft Visual C++ 2013 x64 Additional Runtime"
    "Microsoft Visual C++ 2013 x64 Minimum Runtime"
    "Microsoft Visual C++ 2013 x86 Additional Runtime"
    "Microsoft Visual C++ 2013 x86 Minimum Runtime"
    "Microsoft Visual C++ 2015"
    "Microsoft Visual C++ 2015"
    "Microsoft Visual C++ 2019 X64 Additional Runtime"
    "Microsoft Visual C++ 2019 X64 Minimum Runtime"
    "Microsoft Visual C++ 2019 X86 Additional Runtime"
    "Microsoft Visual C++ 2019 X86 Minimum Runtime"
    "Mozilla Maintenance Service"
    "Office 16 Click-to-Run Extensibility Component"
    "Office 16 Click-to-Run Licensing Component"
    "Office 16 Click-to-Run Localization Component"
    "Realtek Audio COM Components"
    "Realtek Audio Driver"
    "Realtek High Definition Audio Driver"
    "Software Update Wizard (Redist)"
    "Teams Machine-Wide Installer"
    "Windows Firewall Configuration Provider"
);

#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
# Begin Custom Code
#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
If ($DeviceName -eq "CCCJ-FS01") {
    $LogFileDir = $LogFileDir -replace "c",'R:';
    $RecordFileDir = $RecordFileDir -replace "\\\\fs01.criminology.fsu.edu",'R:';
}
If (!$SerialNumber -OR $SerialNumber -eq '       ') {
    If (Test-Path -Path 'C:\CCCJ\ASR\sn.txt' -PathType Leaf) { $SerialNumber = Get-Content -Path 'C:\CCCJ\ASR\sn.txt'; }
}
If ($DeviceName -eq 'EPS-102D-PC01') {
    $RandomNumber = 0;
}
#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
# End Custom Code
#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!


########################################################################################################################################################################################################
# Static Variables 
########################################################################################################################################################################################################
$DataHashTable.Add('SerialNumber', $SerialNumber);
$Win32_ComputerSystem = Get-WmiObject -Class Win32_ComputerSystem;
$CsvFile = "$RecordFileDir\$($SerialNumber).csv";

$DateDir = Get-Date -UFormat "%Y-%B";
$LogFileDir = "$LogFileDir\$DateDir\$($DeviceName)";

$LogFileDate = Get-Date -UFormat "%d-%b-%Y";
$LogFile = "$LogFileDir\$($SerialNumber)_$($DeviceName)_$($LogFileDate)_SelfReport.log";

$StringHasher = [System.Security.Cryptography.HashAlgorithm]::Create('sha256');
$EmailParams.Add('Subject','');
$EmailParams.Add('Body','');
$CustomValues = @{};


########################################################################################################################################################################################################
# Functions
########################################################################################################################################################################################################
Function WriteLog {
	param( [String] $File = $LogFile, [String] $Log, [Object[]] $Data )
    $Date = ((Get-Date -UFormat "%d-%b-%Y_%T") -replace ':', '-');
    If ($LogFile) {
        Switch -WildCard ($Log) {
            "*success*" { Write-Host "[$Date] $Log" -f "Green"; Break; }
            "*ERROR*" { Write-Host "[$Date] $Log" -f "Red"; Break; }
            "*NEW*" { Write-Host "[$Date] $Log" -f "Yellow"; Break; }
            Default { Write-Host "[$Date] $Log" -f "Magenta"; }
        }
        Add-Content $File "[$Date] $Log";
        If ($Data) { ($Data | Out-String).Split("`n") | ForEach-Object { Write-Host $_; Add-Content $File (($_).Trim()); } }
    }
}
Function EmailAlert {
    param( [String] $Subject, [String] $Body )
    $EmailParams.Subject = "$($Subject) - $($DeviceName)";
    $EmailParams.Body = "Device: $($DeviceName)`n`n$($Body)";
    Send-MailMessage @EmailParams;
}
Function GetHRSize {
    param( [INT64] $bytes )
    Process {
        If ( $bytes -gt 1pb ) { "{0:N1}PB" -f ($bytes / 1pb) }
        ElseIf ( $bytes -gt 1tb ) { "{0:N1}TB" -f ($bytes / 1tb) }
        ElseIf ( $bytes -gt 1gb ) { "{0:N1}GB" -f ($bytes / 1gb) }
        ElseIf ( $bytes -gt 1mb ) { "{0:N1}MB" -f ($bytes / 1mb) }
        ElseIf ( $bytes -gt 1kb ) { "{0:N1}KB" -f ($bytes / 1kb) }
        Else   { "{0:N} Bytes" -f $bytes }
    }
}


########################################################################################################################################################################################################
# Create Log Files and Directories
########################################################################################################################################################################################################
Function CheckFilesAndDirectories  {
    param( [Object] $Dir, [Object] $File)
    Try {
        $Dir | ForEach-Object { If (-NOT (Test-Path -Path $_)){ New-Item -ItemType Directory -Path $_; } }
        $File | ForEach-Object { If (-NOT (Test-Path -Path $_ -PathType Leaf)) { New-Item -ItemType File -Path $_ -Force; } }
    } Catch { WriteLog -Log "[ERROR] Error with Directories and Files." -Data $_; }
}
CheckFilesAndDirectories -Dir $LocalFileDir,$LogFileDir,$RecordFileDir -File $LogFile,$StorageAlertLog,$RamAlertLog;


########################################################################################################################################################################################################
# Package Requirements 
########################################################################################################################################################################################################
#WriteLog -Log "Checking Required Packages...";
'NuGet' | ForEach-Object {
        If (-NOT (Get-PackageProvider -ListAvailable -Name $_ -ErrorAction SilentlyContinue)) {
        WriteLog -Log "[LOG] $_ Package not found. Installing...";
        Install-PackageProvider $_ -Confirm:$false -Force:$true;
    } Else {
        $Installed = [String](Get-PackageProvider -ListAvailable -Name $_ | Select-Object -First 1).Version;
        $Latest = [String](Find-PackageProvider -Name $_ | Sort-Object Version -Descending| Select-Object -First 1).version;
        If ([System.Version]$Latest -gt [System.Version]$Installed) {
            WriteLog -Log "[UPDATE] Updating $_...";
            Install-PackageProvider $_ -Confirm:$false -Force:$true;
        }
    }
}


########################################################################################################################################################################################################
# Modules Requirements 
########################################################################################################################################################################################################
#WriteLog -Log "Checking Required Modules...";
If ($Win32_ComputerSystem.Model -eq "Virtual Machine") {
    $RequiredModules = 'SnipeitPS', 'PSWindowsUpdate', 'ActiveDirectory';
} Else { 
    $RequiredModules = 'SnipeitPS', 'DellBIOSProvider', 'ActiveDirectory', 'PSWindowsUpdate'; 
}
$RequiredModules | ForEach-Object {
    Try {
        $Mdle = $_;
        #WriteLog -Log "Checking for $Mdle...";
        If (!(Get-Module -ListAvailable -Name $Mdle)) {
            WriteLog -Log "$Mdle not found. Installing...";
            If ($Mdle -eq 'ActiveDirectory') {
                Add-WindowsCapability –Online –Name “Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0”;
                Install-WindowsFeature RSAT-AD-PowerShell;
            } Else { Install-Module -Name $Mdle -Force; }
        } Else {
            $Latest = [String](Find-Module -Name SnipeitPS | Sort-Object Version -Descending)[0].version;
            $Installed = [String](Get-Module -ListAvailable SnipeitPS | Select-Object -First 1).version;
            If ([System.Version]$Latest -gt [System.Version]$Installed) {
                WriteLog -Log "[UPDATE] Updating $($Mdle)...";
                Update-Module -Name $Mdle -Force;
            }
        }
        Try { Import-Module -Name $Mdle -Force; }
        Catch {
            WriteLog -Log "[ERROR] Unable to Import $($Mdle) Module." -Data $_;
            EmailAlert -Subject "[ERROR] Importing Module" -Body "$($_ | Out-String)";
        }
    } Catch { WriteLog -Log "[ERROR] $($_ | Out-String)"; }
}
#WriteLog -Log "Requirements Installed and Loaded.";


########################################################################################################################################################################################################
# General Device Information
########################################################################################################################################################################################################
#WriteLog -Log "Gathering Device Information...";
$Location = "$(($DeviceName).Split("-")[0])-$(($DeviceName).Split("-")[1])";
$DataHashTable.Add('Location', $Location);
$DataHashTable.Add('DeviceName', $($DeviceName));
$DataHashTable.Add('LastReported', (Get-Date));
$DataHashTable.Add('LastReportedUnix', ([Math]::Round((Get-Date -UFormat %s),0)));
$DataHashTable.Add('Model', $Win32_ComputerSystem.Model);
$DataHashTable.Add('Manufacturer', "$($Win32_ComputerSystem.Manufacturer -replace " Inc.", '')");
$DataHashTable.Add('Bios', $Win32_BIOS.SMBIOSBIOSVersion);


########################################################################################################################################################################################################
# Operating System Information
########################################################################################################################################################################################################
#WriteLog -Log "Gathering Operating System Information...";
$Win32_OperatingSystem = Get-WmiObject -Class Win32_OperatingSystem;
$DataHashTable.Add('OS', ($Win32_OperatingSystem.Name).Split("|")[0]);
$DataHashTable.Add('Build', $Win32_OperatingSystem.Version);
If ($DataHashTable['OS'] -Contains "Server") { $ModelCatID = $Snipe.ServerCatID; }


########################################################################################################################################################################################################
# Bios Information
#################################f###################################################################################################
#WriteLog -Log "Gathering Bios Information...";
If (-NOT ($SerialNumber)) { EmailAlert -Subject "No BIOS Serial Number" -Body ($Win32_BIOS | Out-String); }
Try {
    If ($DataHashTable['Manufacturer'] -eq 'Dell' -AND (Get-Item -Path "DellSmbios:\" -ErrorAction SilentlyContinue)) {
        $BiosChanged = @();
        $Key = Get-Content $KeyFile;
        $OldBiosPwd = Get-Content $OldPwdFile | ConvertTo-SecureString -Key $Key
        $OldBiosPwd = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($OldBiosPwd);
        $OldBiosPwd = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($OldBiosPwd);
        $NewBiosPwd = Get-Content $NewPwdFile | ConvertTo-SecureString -Key $Key
        $NewBiosPwd = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($NewBiosPwd);
        $NewBiosPwd = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($NewBiosPwd);
        If ((Get-Item -Path "DellSmbios:\Security\IsAdminPasswordSet").CurrentValue -eq "True") {
            Try { Set-Item -Path DellSmbios:\Security\AdminPassword "$NewBiosPwd" -Password "$NewBiosPwd" -ErrorAction Stop; } 
            Catch {
                Try { Set-Item -Path DellSmbios:\Security\AdminPassword "$NewBiosPwd" -Password "$OldBiosPwd" -ErrorAction Stop; } 
                Catch { EmailAlert -Subject "Bios Password Change Error" -Body "Unable to change the bios password:`n$($_)"; }
            }
        } Else { Set-Item -Path "DellSmbios:\Security\AdminPassword" "$NewBiosPwd"; }
        Function Set-DellBiosSetting {
            param( [Object[]] $Setting, [String] $Value )
            $CurrentValue = (Get-Item -Path "DellSmbios:\$($Setting.PSChildName)\$($Setting.Attribute)").CurrentValue;
            #Write-Host $CurrentValue
            If ($CurrentValue -ne $Value) {
                $BiosChanged += $Setting.Attribute;
                Try { 
                    Set-Item -Path "DellSmbios:\$($Setting.PSChildName)\$($Setting.Attribute)" -Value $Value -Password $NewBiosPwd;
                    WriteLog -Log "Set Bios Setting: $($Setting.Attribute) to $($Value)."
                } Catch {  WriteLog -Log "Failed to Set Bios Setting." -Data $_; }
            }    
        }

        If ($BiosChanged.Count -gt 0) {
            EmailAlert -Subject "Bios Configurations Changed" -Body "$($BiosChanged -join '`n')";
        }

        ForEach ($Category in (Get-ChildItem -Path "DellSmbios:\").Category) {
            $CategorySettings = Get-ChildItem -Path "DellSmbios:\$($Category)" -WarningAction SilentlyContinue | 
                                   Select-Object Attribute,CurrentValue,PSChildName;
            ForEach ($Setting in $CategorySettings) {
                If ($DataHashTable['BootPathSecurity'] -eq 'UEFI') {
                    Switch ($Setting.Attribute) {
                        "BootList" { $DataHashTable.Add('BootMode', (Get-Item -Path "DellSmbios:\$($Setting.PSChildName)\$($Setting.Attribute)").CurrentValue); }
                        "LegacyOrom" { $DataHashTable.Add('LegacyRoms', (Get-Item -Path "DellSmbios:\$($Setting.PSChildName)\$($Setting.Attribute)").CurrentValue); }
                        "AttemptLegacyBoot" { $DataHashTable.Add('LegacyBoot', (Get-Item -Path "DellSmbios:\$($Setting.PSChildName)\$($Setting.Attribute)").CurrentValue); }
                        "SecureBoot" { $DataHashTable.Add('SecureBoot', (Get-Item -Path "DellSmbios:\$($Setting.PSChildName)\$($Setting.Attribute)").CurrentValue); }
                    }
                }
                #Write-Host "$($Setting.Attribute)  -  $($Setting.CurrentValue)";
                Switch ($Setting.Attribute) {
                    "MemorySpeed" { $MemorySpeed = $Setting.CurrentValue; }
                    "MemoryTechnology" { $MemoryType = $Setting.CurrentValue; }
                    "BootList" { $DataHashTable.Add('BootMode', (Get-Item -Path "DellSmbios:\$($Setting.PSChildName)\$($Setting.Attribute)").CurrentValue); }
                    "LegacyOrom" { $DataHashTable.Add('LegacyRoms', (Get-Item -Path "DellSmbios:\$($Setting.PSChildName)\$($Setting.Attribute)").CurrentValue); }
                    "AttemptLegacyBoot" { $DataHashTable.Add('LegacyBoot', (Get-Item -Path "DellSmbios:\$($Setting.PSChildName)\$($Setting.Attribute)").CurrentValue); }
                    "SecureBoot" { $DataHashTable.Add('SecureBoot', (Get-Item -Path "DellSmbios:\$($Setting.PSChildName)\$($Setting.Attribute)").CurrentValue); }
                    "UefiBootPathSecurity" { $DataHashTable.Add('BootPathSecurity', (Get-Item -Path "DellSmbios:\$($Setting.PSChildName)\$($Setting.Attribute)").CurrentValue); }
                    "EmbNic1" { Set-DellBiosSetting -Setting $Setting -Value "Enabled"; }
                    "SfpNic" { Set-DellBiosSetting -Setting $Setting -Value "EnabledPXE"; }
                    "UefiNwStack" { Set-DellBiosSetting -Setting $Setting -Value "Enabled"; }
                    "SmartErrors" { Set-DellBiosSetting -Setting $Setting -Value "Enabled"; }
                    "TpmSecurity " { Set-DellBiosSetting -Setting $Setting -Value "Enabled"; }
                    "TpmActivation" { Set-DellBiosSetting -Setting $Setting -Value "Enabled"; }
                    "AcPwrRcvry" { Set-DellBiosSetting -Setting $Setting -Value "Last"; }
                    "DeepSleepCtrl" { Set-DellBiosSetting -Setting $Setting -Value "Disabled"; }
                    "WakeOnLan" { Set-DellBiosSetting -Setting $Setting -Value "LanWlan"; }
                    "BlockSleep" { Set-DellBiosSetting -Setting $Setting -Value "Disabled"; }
                    "ChassisIntrusionStatus" { 
                        If ($Setting.CurrentValue -AND $Setting.CurrentValue -ne '' -AND $Setting.CurrentValue -ne "DoorClosed") {
                            EmailAlert -Subject "Chassis Intrustion Detected" -Body "Chassis Status: $($Setting.CurrentValue)";
                            Set-DellBiosSetting -Setting $Setting -Value "TripReset";
                        }
                    }
                    "WirelessLan" { If (-NOT (Get-WmiObject -Class win32_battery)) { Set-DellBiosSetting -Setting $Setting -Value "Disabled"; } }
                    "BluetoothDevice" { If (-NOT (Get-WmiObject -Class win32_battery)) { Set-DellBiosSetting -Setting $Setting -Value "Disabled"; } }
                    "AutoOn" { Set-DellBiosSetting -Setting $Setting -Value "SelectDays"; Break; }
                    "AutoOnHr" { Set-DellBiosSetting -Setting $Setting -Value "7"; Break; }
                    "AutoOnMn" { Set-DellBiosSetting -Setting $Setting -Value "0"; Break; }
                    "AutoOnMon" { If ($DailyPowerOnList -Contains $Location) { Set-DellBiosSetting -Setting $Setting -Value "Enabled"; } Break; }
                    "AutoOnTue" { Set-DellBiosSetting -Setting $Setting -Value "Enabled"; Break; }
                    "AutoOnWed" { If ($DailyPowerOnList -Contains $Location) { Set-DellBiosSetting -Setting $Setting -Value "Enabled"; } Break; }
                    "AutoOnThur" { If ($DailyPowerOnList -Contains $Location) { Set-DellBiosSetting -Setting $Setting -Value "Enabled"; } Break; }
                    "AutoOnFri" { If ($DailyPowerOnList -Contains $Location) { Set-DellBiosSetting -Setting $Setting -Value "Enabled"; } Break; }
                }
            }
        }
    }
} Catch { 
    WriteLog -Log "[ERROR] Issue Configuring Bios"
    WriteLog -Log "$($_ | Out-String)"; 
    EmailAlert -Subject "Error Configuring Bios" -Body "$($_ | Out-String)";
}


########################################################################################################################################################################################################
# Network Adapter Configurations
########################################################################################################################################################################################################
#WriteLog -Log "Gathering Network Adapter Information...";
$MacAddress = @();
$IpAddress = @();
$NetworkAdapters = @();
Get-NetAdapter | Where-Object { $_.Name -NotLike "*bluetooth*" } | ForEach-Object {
    $IfcDesc = $_.InterfaceDescription -replace "\([^\)]+\)",'' -replace '  ',' ';
    $NetworkAdapters += "[$($_.ifIndex)] $($_.LinkSpeed) - $($IfcDesc)";
    $MacAddress += "$($_.MacAddress -replace '-',':') [$($_.ifIndex)]";
    If ($_.Status -eq 'Up') {
        $InterfaceAlias = "$($_.Name)";
        $IpAddress += "$((Get-NetIpAddress | Where-Object { $_.AddressFamily -Like "IPv4" -and $_.InterfaceAlias -eq $InterfaceAlias; }).IPAddress) [$($_.ifIndex)]";
    }
}
$MacAddress = $MacAddress -join "`n";
$IpAddress = $IpAddress -join "`n";
$NetworkAdapters = $NetworkAdapters -join "`n";
$DataHashTable.Add('IpAddress', $IpAddress);
$DataHashTable.Add('MacAddress', $MacAddress);
$DataHashTable.Add('NetworkAdapters', $NetworkAdapters);
Switch ((Get-WmiObject -Class Win32_NetworkAdapterConfiguration | Sort-Object -Property Index | Where-Object { $_.IPAddress } | Select-Object -First 1).DHCPEnabled) {
    "True" { $DataHashTable.Add('DHCP', "Enabled"); Break; }
    "False" { $DataHashTable.Add('DHCP', "Disabled"); Break; }
}


########################################################################################################################################################################################################
# Group Access
########################################################################################################################################################################################################
#WriteLog -Log "Gathering Local Group Information...";
$LocalAdministrators = Get-LocalGroupMember -Group "Administrators";
$DataHashTable.Add('LocalAdmins', ($LocalAdministrators).Name -join "`n");
$RemoteDesktopUsers = Get-LocalGroupMember -Group "Remote Desktop Users";
$DataHashTable.Add('RemoteUsers', ($RemoteDesktopUsers).Name -join "`n");


########################################################################################################################################################################################################
# Uptime
########################################################################################################################################################################################################
#WriteLog -Log "Calculating Uptime...";
$Uptime = "";
$UptimeVal = ((Get-Date)-($Win32_OperatingSystem).ConvertToDateTime($Win32_OperatingSystem.LastBootUpTime));
Switch ($true) {
    ($UptimeVal.Days -gt 0) { $Uptime += "$($UptimeVal.Days)D:"; }
    ($UptimeVal.Hours -gt 0) { $Uptime += "$($UptimeVal.Hours)H:"; }
    ($true) { $Uptime += "$($UptimeVal.Minutes)M:$($UptimeVal.Seconds)S"; $DataHashTable.Add('Uptime', $Uptime); }
}


########################################################################################################################################################################################################
# Software
########################################################################################################################################################################################################
#WriteLog -Log "Gathering Software Information...";
$Apps = @();
$32BitPath = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*";
$64BitPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*";
$Apps += Get-ItemProperty "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" | Select-Object DisplayName,DisplayVersion;
$Apps += Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" | Select-Object DisplayName,DisplayVersion;
$UserDefinedInstallations = @{
    Name = 'DisplayName'; 
    Expression = { 
        If ($_.DisplayName -NotLike "*(User)*") { "$($_.DisplayName) (User)"; } 
        Else { $_.DisplayName -replace "\(USER\)","(User)" } 
    }
}
$Apps += Get-ItemProperty "Registry::\HKEY_CURRENT_USER\$32BitPath" | Select-Object $UserDefinedInstallations,DisplayVersion;
$Apps += Get-ItemProperty "Registry::\HKEY_CURRENT_USER\$64BitPath" | Select-Object $UserDefinedInstallations,DisplayVersion;
$AllProfiles = Get-CimInstance Win32_UserProfile | Select-Object LocalPath, SID, Loaded, Special | Where-Object { $_.SID -Like "S-1-5-21-*" };
$MountedProfiles = $AllProfiles | Where-Object { $_.Loaded -eq $true; }
$UnmountedProfiles = $AllProfiles | Where-Object { $_.Loaded -eq $false; }
$MountedProfiles | ForEach-Object {
    $Apps += Get-ItemProperty -Path "Registry::\HKEY_USERS\$($_.SID)\$($32BitPath)" | Select-Object $UserDefinedInstallations,DisplayVersion;
    $Apps += Get-ItemProperty -Path "Registry::\HKEY_USERS\$($_.SID)\$($64BitPath)" | Select-Object $UserDefinedInstallations,DisplayVersion;
}
$UnmountedProfiles | ForEach-Object {
    $Hive = "$($_.LocalPath)\NTUSER.DAT";
    If (Test-Path $Hive) {
        REG LOAD HKU\temp $Hive | Out-Null;
        $Apps += Get-ItemProperty -Path "Registry::\HKEY_USERS\temp\$($32BitPath)"  | Select-Object $UserDefinedInstallations,DisplayVersion;
        $Apps += Get-ItemProperty -Path "Registry::\HKEY_USERS\temp\$($64BitPath)"  | Select-Object $UserDefinedInstallations,DisplayVersion;
        [GC]::Collect();
        [GC]::WaitForPendingFinalizers();
        REG UNLOAD HKU\temp | Out-Null;
    }
}
$DisplayName = @{ 
    Name = 'DisplayName'; 
    Expression = {
        $Version = $_.DisplayVersion;
        $Name = $_.DisplayName;
        If ($Name -Like "* V$Version*") {
            $NewDisplayName = $Name -replace "V$Version",'';
        } ElseIf ($Name -Like "*$Version*") {
            If ($Name -Like "* - *") { $NewDisplayName = ($Name).Split('-')[0]; } 
            Else { $NewDisplayName = $Name -replace $Version,''; }
        } ElseIf ($Name -Like "*$($Version.SubString(0,$Version.Length-2))") {
            If ($Name -Like "* - *") { $NewDisplayName = ($Name).Split('-')[0]; } 
            Else { $NewDisplayName = $Name -replace "$($Version.SubString(0,$Version.Length-2))",''; }
        } Else { $NewDisplayName = $Name; }
        "$($NewDisplayName)".Trim() -replace '™|®|©','' -replace '\s+', ' ' -replace '[#?\{]','';
    }
}
$NameWithVersion = @{ 
    Name = 'NameWithVersion'; 
    Expression = {
        $Version = $_.DisplayVersion;
        $Name = $_.DisplayName;
        If ($Name -NotLike "*$($Version.SubString(0,$Version.Length-2))*") {
            $NewDisplayName = "$($Name) - $Version";
        } ElseIf ($Name -NotLike "*- $Version") {
            If ($Name -Like "*V$Version") {  $NewDisplayName = "$(($Name -replace "V$Version",'').Trim()) - $Version"; }
            Else { $NewDisplayName = "$(($Name -replace "$Version",'').Trim()) - $Version"; }
        } Else { $NewDisplayName = $Name; }
        "$($NewDisplayName)".Trim() -replace '™|®|©','' -replace '\s+', ' ' -replace '[#?\{]','';
    } 
}
$Software = $Apps | Where-Object { ($_.DisplayName).Length -gt 2; } | Sort-Object DisplayName -Unique | Select-Object $DisplayName,DisplayVersion,$NameWithVersion;
$DataHashTable.Add('SoftwareWithoutVersions', $Software.DisplayName -join ' ; ');
$DataHashTable.Add('SoftwareWithVersions', $Software.NameWithVersion -join ' ; ');
$SoftwareHash = $StringHasher.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($Software.DisplayName -join ' ; '));
$DataHashTable.Add('SoftwareHash', ([System.BitConverter]::ToString($SoftwareHash)).Replace('-', ''));
$InstalledSoftware = (Compare-Object -ReferenceObject $DefaultSoftware -DifferenceObject $Software.DisplayName | Where-Object { $_.SideIndicator -eq "=>"}).InputObject;
$DataHashTable.Add('InstalledSoftware', ($InstalledSoftware -join "`n"));


########################################################################################################################################################################################################
# Drive Configuration Collection 
########################################################################################################################################################################################################
#WriteLog -Log "Gathering Internal and External Drive Information...";
Remove-Variable -Name InternalMedia -ErrorAction SilentlyContinue;
Remove-Variable -Name RemovableMedia -ErrorAction SilentlyContinue;
$DiskDrives = Get-WmiObject Win32_DiskDrive -Property * | Sort-Object DeviceID;
$DiskVolumes = Get-Volume | Sort-Object Index;
$PhysicalDisks = Get-PhysicalDisk;
$InternalDisks = @();
$InternalMedia = @();
$RemovableMedia = @();
$UnhealthyDisks = @();
$LowSpaceDrives = @();
ForEach ($Disk in $DiskDrives) {
    Remove-Variable -Name DiskInfo -ErrorAction SilentlyContinue;
    Remove-Variable -Name DriveType -ErrorAction SilentlyContinue;
    $PhysicalDisk = ($PhysicalDisks | Where-Object { $_.DeviceID -eq (($Disk.DeviceID).substring((($Disk.DeviceID).Length)-1)) });
    Switch ($PhysicalDisk.MediaType) {
        'Unspecified' { $Disk.MediaType = 'USB'; Break; }
        'External hard disk media' { $Disk.MediaType = 'HDD'; Break; }
        default { $Disk.MediaType = $PhysicalDisk.MediaType; }
    }
    If ($PhysicalDisk.BusType -eq $Disk.MediaType) { $DiskType = $PhysicalDisk.BusType; }
    If ($PhysicalDisk.HealthStatus -ne 'Healthy') { $UnhealthyDisks += $PhysicalDisk; }
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
            $DriveType = @($DiskVolume.DriveType,'Removable')[($PhysicalDisk.BusType -eq 'USB')]
            $DiskInfo += "--------- $VolumeData`n";
            If ($DriveType -ne 'Removable' -AND (($DiskVolume.SizeRemaining / $DiskVolume.Size) -lt .1)) {
                $LowSpaceDrives += [PSCustomObject]@{
                    Drive = $DiskVolume.DriveLetter;
                    SpaceAvailable = $DiskVolume.SizeRemaining;
                    TotalSize = $DiskVolume.Size;
                } 
            }
        }
    }
    Switch ($DriveType) {
        "Removable" { $RemovableMedia += $DiskInfo.Trim(); Break; }
        default { 
            $InternalDisks += "[$($Disk.MediaType)] $($Disk.Model) ($(GetHRSize $Disk.size))";
            $InternalMedia += $DiskInfo.Trim(); Break; 
        }
    }
}
$DataHashTable.Add('Drives', $InternalDisks -join "`n");
$DataHashTable.Add('InternalMedia', $InternalMedia -join "`n");
$DataHashTable.Add('RemovableMedia', $RemovableMedia -join "`n");
If ($LowSpaceDrives.Count -gt 0) { 
    $LastStorageAlert = Get-Content $StorageAlertLog | ConvertFrom-Json;
    $TimeSinceLastStorageAlert = New-TimeSpan -Start (Get-Date -Date $LastStorageAlert.Last_Notified) -End (Get-Date);
    $StorageNotification = [PSCustomObject]@{
        'Drives' = $LowSpaceDrives;
        'Last_Notified' = (Get-Date).DateTime;
    }
    If (!$TimeSinceLastStorageAlert -OR (($StorageNotification.Drives).Drive | Out-String) -ne (($LastStorageAlert.Drives).Drive | Out-String) -OR $TimeSinceLastStorageAlert.TotalDays -gt 30) {
        EmailAlert -Subject "Drive Usage Very High" -Body "$($StorageNotification | Format-List | Out-String)`n$InternalMedia`n$RemovableMedia"; 
        Set-Content -Path $StorageAlertLog -Value ($StorageNotification | ConvertTo-Json)
    }
}
If ($UnhealthyDisks.Count -gt 0) { EmailAlert -Subject "Unhealthy Drive(s) Detected" -Body "$($UnhealthyDisks | Format-List | Out-String)"; }


########################################################################################################################################################################################################
# Logged In Users
########################################################################################################################################################################################################
#WriteLog -Log "Gathering Active User Information...";
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
   EmailAlert -Subject "[ERROR] Unable to Collect User Information." -Body "$($_ | Out-String)";
}


########################################################################################################################################################################################################
# Webcam IdentIfication
########################################################################################################################################################################################################
#WriteLog -Log "Gathering Webcam Information...";
$Webcam = Get-PnpDevice | Where-Object  { $_.Status -ne 'Unknown' -AND $_.Class -eq "Image" -AND $_.FriendlyName -Like "*webcam*" } | Get-PnpDeviceProperty;
$WebcamGuid = ($Webcam | Where-Object { $_.KeyName -eq "DEVPKEY_Device_ContainerId" }).Data;
$Webcam = ($Webcam | Where-Object { $_.KeyName -eq "DEVPKEY_Device_FriendlyName" }).Data;
If ($Webcam.Count -gt 1) { 
    EmailAlert -Subject "Multiple Webcams Assigned to Location" -Body "$($Webcam | Out-String)"; 
}
Switch ($true) {
    ($null -ne $Webcam) { $DataHashTable.Add('Webcam', "$($Webcam)"); Break; } 
    default { $DataHashTable.Add('Webcam', ''); }
}


########################################################################################################################################################################################################
# GPU IdentIfication
########################################################################################################################################################################################################
#WriteLog -Log "Gathering Graphics Information...";
$GraphicsCard = (Get-PnpDevice | Where-Object {$_.Class -eq "Display" -AND $_.Status -eq 'OK'} | 
                    Get-PnpDeviceProperty | Where-Object { $_.Keyname -eq "DEVPKEY_NAME" } | 
                    Where-Object { $_.Data -ne "Microsoft Remote Display Adapter" } | 
                    Sort-Object -Property Data).Data -join "`n";
Switch ($true) {
    ($null -ne $GraphicsCard) { $DataHashTable.Add('Graphics', "$($GraphicsCard)"); Break; } 
    default { $DataHashTable.Add('Graphics', ''); }
}


########################################################################################################################################################################################################
# RAM/Memory
########################################################################################################################################################################################################
#WriteLog -Log "Gathering Memory Information...";
$Memory = Get-WmiObject -Class Win32_PhysicalMemory | Select-Object * -First 1;
$MemoryVoltage = $Memory.ConfiguredVoltage;
If (-NOT ($MemorySpeed)) { $MemorySpeed = "$($Memory.Speed)MHz"; }
If (-NOT ($MemoryType)) { 
    Switch ($MemoryVoltage) { 
        '1200' { $MemoryType = "DDR4"; } 
        '1500' { $MemoryType = "DDR3"; } 
        default { $MemoryType = ''; } 
    }
}
$Memory = Get-WmiObject -Class Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum;

$MemoryAvailable = [math]::Round(($Win32_OperatingSystem.FreePhysicalMemory / 1MB),1);
$MemoryUsed = [math]::Round((($Memory.Sum / 1gb)-($Win32_OperatingSystem.FreePhysicalMemory / 1MB)),1);
$MemoryInstalled = "$($Memory.Sum / 1gb)GB";

$DataHashTable.Add('RAM', "$($MemoryUsed)/$($MemoryInstalled) [$($Memory.Count)] $($MemorySpeed) $($MemoryType)");
$DataHashTable.Add('RAM_Installed', "$($MemoryInstalled) [$($Memory.Count)]");

If ([int]$MemoryAvailable -lt 1 -AND $DataHashTable['Model'] -ne "Virtual Machine") {
    $LastRamAlert = Get-Content $RamAlertLog | ConvertFrom-Json;
    $TimeSinceLastRamAlert = New-TimeSpan -Start (Get-Date -Date $LastRamAlert.Last_Notified) -End (Get-Date);
    If (!$TimeSinceLastRamAlert -OR $TimeSinceLastRamAlert.TotalDays -gt 30) {
        $RamNotification = [PSCustomObject]@{
            'RAM_Installed' = "$($MemoryInstalled) [$($Memory.Count)]";
            'RAM_Available' = "$($MemoryAvailable)Gb";
            'Last_Notified' = (Get-Date).DateTime;
            'Previous_Notification' = $LastRamAlert.Last_Notified;
        }
        EmailAlert -Subject "Low RAM Availability" -Body "$($RamNotification | Format-List | Out-String)";
        Set-Content -Path $RamAlertLog -Value ($RamNotification | ConvertTo-Json)
    }
}


########################################################################################################################################################################################################
# CPU/Processor
########################################################################################################################################################################################################
#WriteLog -Log "Gathering Processor Information...";
$Win32_Processor = (Get-WmiObject Win32_Processor | Select-Object *);
If ($Win32_Processor.Count -gt 1) { 
    $Win32_Processor = $Win32_Processor[0]; 
    $Win32_Processor.Name = "[2] $($Win32_Processor.Name)"; 
}
$AssetProcessor = ($Win32_Processor.Name -replace '\(TM\)|\(R\)','');
$AssetProcessor = ($AssetProcessor -replace '@',"$($Win32_Processor.NumberOfCores)c/$($Win32_Processor.NumberOfLogicalProcessors)t");
$AssetProcessor = ($AssetProcessor -replace '  | 0 ',' ');
$DataHashTable.Add('CPU', $AssetProcessor);


########################################################################################################################################################################################################
# Applied Updates
########################################################################################################################################################################################################
#WriteLog -Log "Fetching Applied Updates...";
$AppliedUpdates = "";
Get-WmiObject -Class Win32_QuickfixEngineering | Where-Object { $_.InstalledOn -ge ((Get-Date).AddMonths(-12)); } | Sort-Object InstalledOn -Descending -ErrorAction SilentlyContinue | ForEach-Object {
    If ($AppliedUpdates.Length -gt 0) { $AppliedUpdates +="`n"; }
    $AppliedUpdates += "$($_.HotFixID)   $(([string]($_.InstalledOn)).Split(' ')[0])";
}
$DataHashTable.Add('AppliedUpdates', $AppliedUpdates);


########################################################################################################################################################################################################
# Load Record If Exists
########################################################################################################################################################################################################
#WriteLog -Log "Loading Existing Record...";
If (Test-Path -Path $CsvFile -PathType Leaf) {
    $Record = Import-Csv -Path $CsvFile;
}


########################################################################################################################################################################################################
# Fetch Dell Warranty Data
########################################################################################################################################################################################################
#WriteLog -Log "Checking Dell Warranty Information...";
$DellUri = "https://apigtwb2c.us.dell.com/PROD/sbil/eapi/v5/asset-entitlements";
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
                $DellApiResponse = Invoke-RestMethod -Uri $DellUri -Headers $DellApiHeaders -Body $DellApiParams -Method Get -ContentType "application/json" -ea 0;
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


########################################################################################################################################################################################################
# Update SnipeIT 
########################################################################################################################################################################################################
#WriteLog -Log "Checking in to SnipeIT...";
Start-Sleep -Seconds $RandomNumber;
Connect-SnipeitPS -URL $Snipe.Url -apiKey $Snipe.Token;
$SnipeAsset = Get-SnipeItAsset -asset_serial $DataHashTable['SerialNumber'];
If ($SnipeAsset.StatusCode -eq 'InternalServerError') {
    $SnipeAsset = Get-SnipeItAsset -asset_serial $DataHashTable['SerialNumber'];
    If ($SnipeAsset.StatusCode -eq 'InternalServerError') {
        EmailAlert -Subject "Error searching SnipeIT" -Body "(Duplicate Check)`n$($SnipeAsset | Format-List | Out-String)";
        Exit 0;
    }
}
$CustomValues.Add('purchase_date', $DataHashTable['Purchased']);
$CustomValues.Add('warranty_months', $DataHashTable['WarrantyMonths']);
$CustomValues.Add('_snipeit_mac_address_1', $DataHashTable['MacAddress']);
$CustomValues.Add('_snipeit_cpu_2', $DataHashTable['CPU']);
$CustomValues.Add('_snipeit_ram_3', $DataHashTable['RAM']);
$CustomValues.Add('_snipeit_operating_system_5', $DataHashTable['OS']);
$CustomValues.Add('_snipeit_ip_address_9', $DataHashTable['IpAddress']);
$CustomValues.Add('_snipeit_bios_11', $DataHashTable['Bios']);
$CustomValues.Add('_snipeit_last_reported_12', (Get-Date -UFormat "%d-%b-%Y %T"));
$CustomValues.Add('_snipeit_graphics_13', $DataHashTable['Graphics']);
$CustomValues.Add('_snipeit_boot_drive_15', $DataHashTable['BootDrive']);
$CustomValues.Add('_snipeit_internal_media_16', $DataHashTable['InternalMedia']);
$CustomValues.Add('_snipeit_external_media_17', $DataHashTable['RemovableMedia']);
$CustomValues.Add('_snipeit_installed_software_18', $DataHashTable['InstalledSoftware']);
$CustomValues.Add('_snipeit_remote_desktop_users_19', $DataHashTable['RemoteUsers']);
$CustomValues.Add('_snipeit_applied_updates_22', $DataHashTable['AppliedUpdates']);
$CustomValues.Add('_snipeit_network_adapters_24', $DataHashTable['NetworkAdapters']);
$CustomValues.Add('_snipeit_age_25', $DataHashTable['Age']);
$NextAuditDate = Get-Date;
If ($NextAuditDate.Month -ne 1) {
    $NextAuditDate = New-Object DateTime(($NextAuditDate.Year+1), 1, [DateTime]::DaysInMonth($NextAuditDate.Year, $NextAuditDate.Month))
    $Diff = ([int] [DayOfWeek]::Friday) - ([int]$lastDay.DayOfWeek);
    $NextAuditDate = @((Get-Date -Date $NextAuditDate.AddDays($Diff)),(Get-Date -Date $NextAuditDate.AddDays(- (7-$Diff))))[($Diff -ge 0)];
    $CustomValues.Add('next_audit_date', ($NextAuditDate | Get-Date -UFormat "%Y-%m-%d"));
}
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
        $SnipeAsset = New-SnipeItAsset -name $DataHashTable['DeviceName'] -status_id 5 -model_id $Model.id -serial $DataHashTable['SerialNumber'] -asset_tag $DataHashTable['SerialNumber'] -customfields $CustomValues;
        WriteLog -Log "[SnipeIT] Created a new Asset in SnipeIT.";
    } Catch { WriteLog -Log "[SnipeIT] [ERROR] Unable to Create new Asset." -Data $_; }
} ElseIf ($SnipeAsset.Count -gt 1) {
    WriteLog -Log "[ERROR] Multiple Assets with Identical Serial Numbers Found in SnipeIT.";
    EmailAlert -Subject "[Inventory Discrepancy] Multiple Assets with Identical Serial Numbers Found" -Body "Asset Name: $($DeviceName)`n`n$($SnipeAsset | Format-List | Out-String)";
} Else {
    # Check Assigned User and Remote Users
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
    # Check Asset Tag
    If ($SnipeAsset.asset_tag) { $DataHashTable.Add('AssetTag', $SnipeAsset.asset_tag); }
    # Check Location
    If ($DataHashTable['Location'] -NotLike "CCCJ-*") {
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
            EmailAlert -Subject "[Inventory Discrepancy] Update Asset Location in SnipeIT" -Body "SnipeIT Location: $($SnipeLocation)`nAsset Name: $($DeviceName)`n$($SnipeAsset | Format-List | Out-String)";
        }
    }
    # Check Webcam
    If ($Webcam -AND !($Webcam.Count -gt 1)) {
        $SnipeWebcam = Get-SnipeItAsset -location_id $SnipeAsset.location.id | Where-Object { $Webcam -like $_.model.name };
        If ($SnipeWebcam) { 
            If (!$SnipeWebcam.custom_fields.GUID.value -OR $SnipeWebcam.custom_fields.GUID.value -eq '') {
                Set-SnipeitAsset -id $SnipeWebcam.id -customfields @{ _snipeit_guid_25=$WebcamGuid; }
            } ElseIf ($SnipeWebcam.custom_fields.GUID.value -ne $WebcamGuid) {
                WriteLog -Log "Webcam GUIDs Do not Match!";
                EmailAlert -Subject "Webcam GUIDs Do Not Match!" -Body "Assigned Webcam GUID: $($SnipeWebcam.custom_fields.GUID.value)`nSeen Webcam GUID: $($WebcamGuid)"
            }
            If (!$SnipeWebcam.custom_fields.Host.value -OR $SnipeWebcam.custom_fields.Host.value -eq '') {
                Set-SnipeitAsset -id $SnipeWebcam.id -customfields @{ _snipeit_host_26=$DeviceName; }
            } If ($SnipeWebcam.custom_fields.Host.value -ne $DeviceName) {
                WriteLog -Log "Assigned Hosts for the Webcam do not Match!";
                EmailAlert -Subject "Webcam Hosts Does Not Match!" -Body "Assigned Webcam Host: $($SnipeWebcam.custom_fields.Host.value)`nSeen Webcam Host: $($DeviceName)"
            }
        } Else {
            $SnipeWebcam = Get-SnipeItAsset -Search $WebcamGuid;
            If ($SnipeWebcam) { EmailAlert -Subject "Webcam not Assigned to Location" -Body "$($SnipeWebcam | Out-String)"; }
        }
    }
    # Check License Software
    Try {
        #WriteLog -Log "Checking Installed Software against Inventory...";
        $AssetData = $($SnipeAsset | Select-Object @{N='Name';E={$_.name}},@{N='AssetTag';E={$_.asset_tag}},@{N='Serial';E={$_.serial}},@{N='AssignedTo'; E={$_.assigned_to.username}} | Format-List | Out-String);
        If ($InstalledSoftware) {
            $InstalledSoftware | ForEach-Object {
                $SW = ($_).Split('-')[0].Trim();
                Switch -Wildcard ($SW) {
                    'Stata*' {
                        If (Get-ChildItem -Path "C:\Program Files\$($SW -replace ' ','')*\STATA.LIC") {
                            $StataLicense = (Get-ChildItem -Path "C:\Program Files\$($SW -replace ' ','')*\STATA.LIC" | Get-Content).Split('!');
                        } ElseIf (Get-ChildItem -Path "C:\Program Files (x86)\$($SW -replace ' ','')*\STATA.LIC") {
                            $StataLicense = (Get-ChildItem -Path "C:\Program Files (x86)\$($SW -replace ' ','')*\STATA.LIC" | Get-Content).Split('!');
                        }
                        If (!$StataLicense) {
                            EmailAlert -Subject "[Software Discrepancy] Unlicensed Stata Install" -Body "No Stata license is configured for install on this PC. $SW`n`n$($InstalledSoftware | Out-String)";
                        } Else {
                            $LocalStataSerial = $StataLicense[0];
                            $AssignedLicense = Get-SnipeItLicense -asset_id $SnipeAsset.id | Where-Object { $_.name -Like 'Stata*' } | Select-Object id,name,product_key;
                            If (!$AssignedLicense) {
                                $SnipeItStataLicense = Get-SnipeItLicense | Where-Object { $_.name -Like "Stata*" -AND ([BigInt]($_.product_key.Split("`n")[0]) -eq [BigInt]$LocalStataSerial) };
                                If ($SnipeItStataLicense) {
                                    $OpenStataSeats = Get-SnipeItLicenseSeat -id $SnipeItStataLicense.id | Where-Object { !$_.assigned_asset } | Sort-Object -Property name;
                                    If (!$OpenStataSeats) { EmailAlert -Subject "Stata License Error" -Body "The Stata license assigned has no open seats.`n`nInstalled:`n$($StataLicense | Out-String)"; }
                                    Else { Set-SnipeItLicenseSeat -id $SnipeItStataLicense.id -seat_id $OpenStataSeats[0].id -assigned_to $SnipeAsset.assigned_to.id -asset_id $SnipeAsset.id; }
                                } Else { EmailAlert -Subject "Unknown Stata License" -Body "A Stata license was seen that does not match anything in the inventory.`n`nInstalled:`n$($StataLicense)"; }
                            } Else {
                                $SnipeItStataSerial = $AssignedLicense.product_key.Split("`n")[0];
                                If ($AssignedLicense.Count -gt 1) {
                                    EmailAlert -Subject "[Inventory Discrepancy] Multiple Stata Licenses Assigned" -Body "This asset is using multiple license seats.`n`nInstalled:`n$($StataLicense)`n`nAssigned:`n$($AssignedLicense | Out-String)";
                                } ElseIf ([BigInt]$SnipeItStataSerial -ne [BigInt]$LocalStataSerial) {
                                    EmailAlert -Subject "[Inventory Discrepancy] Mismatched Stata License" -Body "A Stata license assigned is not the same license that is instaled.`n`nInstalled:`n$($StataLicense)`n`nAssigned:`n$($AssignedLicense | Out-String)";
                                } Else {
                                    $LicenseSeat = Get-SnipeItLicenseSeat -id $AssignedLicense.id | Where-Object { $_.assigned_asset.id -eq $SnipeAsset.id };
                                    If ($SnipeAsset.assigned_to.username -AND !$LicenseSeat.assigned_user) {
                                        Set-SnipeItLicenseSeat -id $AssignedLicense.id -seat_id $LicenseSeat.id -assigned_to $SnipeAsset.assigned_to.id -asset_id $SnipeAsset.id;
                                    } ElseIf ($SnipeAsset.assigned_to.username -AND ($LicenseSeat.assigned_user.id -ne $SnipeAsset.assigned_to.id)) {
                                        EmailAlert -Subject "Stata License Error" -Body "The Stata license assigned user does not match the Asset's assigned user.`nLicense: $($LicenseSeat.assigned_user)`nAsset: $($SnipeAsset.assigned_to) $($AssetData)";
                                    }
                                }
                            }
                        }
                    }
                    'Stat/Transfer*' {
                        $AssignedLicense = Get-SnipeItLicense -asset_id $SnipeAsset.id | Where-Object { $_.name -Like 'Stat/Transfer*' } | Select-Object id,name,product_key;
                        If (!$AssignedLicense) {
                            $StatTransferLicense = Get-SnipeItLicense | Where-Object { $_.name -Like "Stat/Transfer*" };
                            If ($StatTransferLicense) {
                                $OpenStatTransferSeats = Get-SnipeItLicenseSeat -id $StatTransferLicense.id | Where-Object { !$_.assigned_asset } | Sort-Object -Property name;
                                If (!$OpenStatTransferSeats) { EmailAlert -Subject "[Inventory Discrepancy] Stat/Transfer License Error" -Body "The Stat/Transfer license assigned has no open seats."; }
                                Else { Set-SnipeItLicenseSeat -id $StatTransferLicense.id -seat_id $OpenStatTransferSeats[0].id -assigned_to $SnipeAsset.assigned_to.id -asset_id $SnipeAsset.id; }
                            } Else { EmailAlert -Subject "[Inventory Discrepancy] Unknown Stat/Transfer License" -Body "Cannot find a Stat/Transfer license in Snipe IT."; }
                        } Else {
                            If ($AssignedLicense.Count -gt 1) {
                                $EmailSubject = "[Inventory Discrepancy] Multiple Stat/Transfer Licenses Assigned";
                                $EmailBody = "This asset is using multiple license seats.`n`nInstalled:`n$($StatTransferLicense)`n`nAssigned:`n$($AssignedLicense | Out-String)";
                                EmailAlert -Subject $EmailSubject -Body $EmailBody;
                            } Else {
                                $LicenseSeat = Get-SnipeItLicenseSeat -id $AssignedLicense.id | Where-Object { $_.assigned_asset.id -eq $SnipeAsset.id; };
                                If ($SnipeAsset.assigned_to.username -AND !$LicenseSeat.assigned_user) {
                                    Set-SnipeItLicenseSeat -id $AssignedLicense.id -seat_id $LicenseSeat.id -assigned_to $SnipeAsset.assigned_to.id -asset_id $SnipeAsset.id;
                                } ElseIf ($SnipeAsset.assigned_to.username -AND  $LicenseSeat.assigned_user.id -ne $SnipeAsset.assigned_to.id) {
                                    $EmailSubject = "Stat/Transfer License Error";;
                                    $EmailBody = "The Stat/Transfer license assigned user does not match the Asset's assigned user.`nLicense: $($LicenseSeat.assigned_user)`nAsset: $($SnipeAsset.assigned_to) $($AssetData)";
                                    EmailAlert -Subject $EmailSubject -Body $EmailBody;
                                }
                            }
                        }
                    }
                    'HLM*' {
                        $AssignedLicense = Get-SnipeItLicense -asset_id $SnipeAsset.id | Where-Object { $_.name -Like "HLM*"; } | Select-Object id,name,product_key;
                        If (!$AssignedLicense) {

                        } Else {
                            $LicenseSeat = Get-SnipeItLicenseSeat -id $AssignedLicense.id | Where-Object { $_.assigned_asset.id -eq $SnipeAsset.id };
                            If ($SnipeAsset.assigned_to.username -AND !$LicenseSeat.assigned_user) {
                                Set-SnipeItLicenseSeat -id $AssignedLicense.id -seat_id $LicenseSeat.id -assigned_to $SnipeAsset.assigned_to.id -asset_id $SnipeAsset.id;
                            } ElseIf ($SnipeAsset.assigned_to.username -AND  $LicenseSeat.assigned_user.id -ne $SnipeAsset.assigned_to.id) {
                                $EmailSubject = "HLM License Error";
                                $EmailBody = "The HLM license assigned user does not match the Asset's assigned user.`nLicense: $($LicenseSeat.assigned_user)`nAsset: $($SnipeAsset.assigned_to) $($AssetData)";
                                EmailAlert -Subject $EmailSubject -Body $EmailBody;
                            }
                        }
                    }
                    'MPlus*' {
                        $AssignedLicense = Get-SnipeItLicense -asset_id $SnipeAsset.id | Where-Object { $_.name -Like "MPlus*"; };
                        If (!$AssignedLicense) {
                            EmailAlert -Subject "[Inventory Discrepancy] Unassigned Mplus License" -Body "Mplus is installed, but a license is not assigned in the inventory system.$($AssetData)";
                        }
                    }
                    'SAS*' {
                        $AssignedLicense = Get-SnipeItLicense -asset_id $SnipeAsset.id | Where-Object { $_.name -Like "SAS*"; } | Select-Object id,name,product_key;
                        If (!$AssignedLicense) {
                            $SnipeItSasLicense = Get-SnipeItLicense | Where-Object { $_.name -Like "SAS*"; };
                            If ($SnipeItSasLicense) {
                                $OpenSasSeats = Get-SnipeItLicenseSeat -id $SnipeItSasLicense.id | Where-Object { !$_.assigned_asset } | Sort-Object -Property name;
                                If (!$OpenSasSeats) { EmailAlert -Subject "[Inventory Discrepancy] SAS License Error" -Body "The SAS license assigned has no open seats."; } 
                                Else { Set-SnipeItLicenseSeat -id $SnipeItSasLicense.id -seat_id $OpenSasSeats[0].id -assigned_to $SnipeAsset.assigned_to.id -asset_id $SnipeAsset.id; }
                            } Else { EmailAlert -Subject "[Inventory Discrepancy] Unknown SAS License" -Body "A SAS license cannot be found in the inventory."; }
                        } Else {
                            If ($AssignedLicense.Count -gt 1) {
                                $EmailSubject = "[Inventory Discrepancy] Multiple SAS Licenses Assigned";
                                $EmailBody = "This asset is using multiple license seats.`n`nInstalled:`n$($StatTransferLicense)`n`nAssigned:`n$($AssignedLicense| Out-String)";
                                EmailAlert -Subject $EmailSubject -Body $EmailBody;
                            } Else {
                                $LicenseSeat = Get-SnipeItLicenseSeat -id $AssignedLicense.id | Where-Object { $_.assigned_asset.id -eq $SnipeAsset.id; };
                                If ($SnipeAsset.assigned_to.username -AND !$LicenseSeat.assigned_user) {
                                    Set-SnipeItLicenseSeat -id $AssignedLicense.id -seat_id $LicenseSeat.id -assigned_to $SnipeAsset.assigned_to.id -asset_id $SnipeAsset.id;
                                } ElseIf ($SnipeAsset.assigned_to.username -AND  $LicenseSeat.assigned_user.id -ne $SnipeAsset.assigned_to.id) {
                                    $EmailSubject = "SAS License Error";
                                    $EmailBody = "The SAS license assigned user does not match the Asset's assigned user.`nLicense: $($LicenseSeat.assigned_user)`nAsset: $($SnipeAsset.assigned_to) $($AssetData)";
                                    EmailAlert -Subject $EmailSubject -Body $EmailBody;
                                }
                            }
                        }
                    }
                    'IBM SPSS*' {
                        $AssignedLicense = Get-SnipeItLicense -asset_id $SnipeAsset.id | Where-Object { $_.name -Like "IBM SPSS*"; } | Select-Object id,name,product_key;
                        If (!$AssignedLicense) {
                            $SnipeItSpssLicense = Get-SnipeItLicense | Where-Object { $_.name -Like "IBM SPSS*"; };
                            If ($SnipeItSpssLicense) {
                                $OpenSpssSeats = Get-SnipeItLicenseSeat -id $SnipeItSpssLicense.id | Where-Object { !$_.assigned_asset } | Sort-Object -Property name;
                                If (!$OpenSpssSeats) { EmailAlert -Subject "[Inventory Discrepancy] SPSS License Error" -Body "The SPSS license assigned has no open seats."; } 
                                Else { $AssignedLicense = Set-SnipeItLicenseSeat -id $SnipeItSpssLicense.id -seat_id $OpenSpssSeats[0].id -asset_id $SnipeAsset.id; }
                            } Else { EmailAlert -Subject "[Inventory Discrepancy] Unknown SPSS License" -Body "An SPSS license cannot be found in the inventory."; }
                        } Else {
                            $LicenseSeat = Get-SnipeItLicenseSeat -id $AssignedLicense.id | Where-Object { $_.assigned_asset.id -eq $SnipeAsset.id };
                            If ($SnipeAsset.assigned_to.username -AND !$LicenseSeat.assigned_user) {
                                Set-SnipeItLicenseSeat -id $AssignedLicense.id -seat_id $LicenseSeat.id -assigned_to $SnipeAsset.assigned_to.id -asset_id $SnipeAsset.id;
                            } ElseIf ($SnipeAsset.assigned_to.username -AND  $LicenseSeat.assigned_user.id -ne $SnipeAsset.assigned_to.id) {
                                $EmailSubject = "SPSS License Error";
                                $EmailBody = "The SPSS license assigned user does not match the Asset's assigned user.`nLicense: $($LicenseSeat.assigned_user)`nAsset: $($SnipeAsset.assigned_to) $($AssetData)";
                                EmailAlert -Subject $EmailSubject -Body $EmailBody;
                            }
                        }
                    }
                    'EndNote*' {
                        $AssignedLicense = Get-SnipeItLicense -asset_id $SnipeAsset.id | Where-Object { $_.name -Like 'EndNote*' } | Select-Object id,name,product_key;
                        If (!$AssignedLicense) {

                        } Else {
                            $LicenseSeat = Get-SnipeItLicenseSeat -id $AssignedLicense.id | Where-Object { $_.assigned_asset.id -eq $SnipeAsset.id };
                            If ($SnipeAsset.assigned_to.username -AND !$LicenseSeat.assigned_user) {
                                Set-SnipeItLicenseSeat -id $AssignedLicense.id -seat_id $LicenseSeat.id -assigned_to $SnipeAsset.assigned_to.id -asset_id $SnipeAsset.id;
                            } ElseIf ($SnipeAsset.assigned_to.username -AND  $LicenseSeat.assigned_user.id -ne $SnipeAsset.assigned_to.id) {
                                $EmailSubject = "SPSS License Error";
                                $EmailBody = "The SPSS license assigned user does not match the Asset's assigned user.`nLicense: $($LicenseSeat.assigned_user)`nAsset: $($SnipeAsset.assigned_to) $($AssetData)";
                                EmailAlert -Subject $EmailSubject -Body $EmailBody;
                            }
                        }
                    }
                    'ArcGIS*' {
                        $AssignedLicense = Get-SnipeItLicense -asset_id $SnipeAsset.id | Where-Object { $_.name -Like 'ArcGIS*' } | Select-Object id,name,product_key;
                        If (!$AssignedLicense) {
                            
                        } Else {
                            $LicenseSeat = Get-SnipeItLicenseSeat -id $AssignedLicense.id | Where-Object { $_.assigned_asset.id -eq $SnipeAsset.id };
                            If ($SnipeAsset.assigned_to.username -AND !$LicenseSeat.assigned_user) {
                                Set-SnipeItLicenseSeat -id $AssignedLicense.id -seat_id $LicenseSeat.id -assigned_to $SnipeAsset.assigned_to.id -asset_id $SnipeAsset.id;
                            } ElseIf ($SnipeAsset.assigned_to.username -AND  $LicenseSeat.assigned_user.id -ne $SnipeAsset.assigned_to.id) {
                                $EmailSubject = "ArcGIS License Error";
                                $EmailBody = "The ArcGIS license assigned user does not match the Asset's assigned user.`nLicense: $($LicenseSeat.assigned_user)`nAsset: $($SnipeAsset.assigned_to) $($AssetData)";
                                EmailAlert -Subject $EmailSubject -Body $EmailBody;
                            }
                        }
                    }
                }
            }
        }
    } Catch { EmailAlert -Subject "Software License Inventory Check Error" -Body "$($_ | Out-String)"; }
    # Audit Assigned Licenses
    $AssignedLicenses = Get-SnipeItLicense -asset_id $SnipeAsset.id;
    If ($AssignedLicenses) {
        
    }
    # Update Asset
    Try {
        $UpdatedAsset = Set-SnipeItAsset -name $DataHashTable['DeviceName'] -id $SnipeAsset.id -status_id $Snipe.DefStatusID -customfields $CustomValues;
        WriteLog -Log "[SnipeIT] Updated an Asset in SnipeIT." -Data $UpdatedAsset;
        $SnipeAsset = $UpdatedAsset;
    } Catch { WriteLog -Log "[ERROR] Unable to Update SnipeIT Asset." -Data $_; }
}

# Check for Duplicate objects in SnipeIT
$DuplicateNames = Get-SnipeItAsset -Search $DataHashTable['DeviceName'] | Where-Object { $_.serial -ne $DataHashTable['SerialNumber'] -AND $_.assigned_to.name -NotLike "$($DeviceName)*" };
If ($DuplicateNames.StatusCode -eq 'InternalServerError') {
    EmailAlert -Subject "Error searching SnipeIT" -Body "(Duplicate Check)`n$($DuplicateNames | Format-List | Out-String)";
} ElseIf ($DuplicateNames.Count -gt 0) {
    ForEach ($Duplicate in $DuplicateNames) {
        $UpdatedAsset = Set-SnipeItAsset -name '' -id $Duplicate.id -customfields { _snipeit_ip_address_9 = ''; };
    }
    WriteLog -Log "[ERROR] Duplicate Information Found in SnipeIT.";
    EmailAlert -Subject "Multiple Assets with Identical Data Found. Removed data from inventory object." -Body "Asset Name: $($DeviceName)`n`n$($DuplicateNames | Format-List | Out-String)";
}


########################################################################################################################################################################################################
# Save New Record If Not Exists
##########################################################################
$DataObject = [PSCustomObject]$DataHashTable;
If (!$Record) {
    Try {
        $DataObject | Export-Csv -Path $CsvFile;
        WriteLog -Log "[REPORT] [NEW] $($DataObject.DeviceName) Self-Reported" -Data $DataObject;
        EmailAlert -Subject "New Asset Self-Reported" -Body ($DataObject | Format-List | Out-String);
    } Catch { WriteLog -Log "[ERROR] Error Saving Excel File.";  WriteLog -Log "[ERROR] $($_ | Out-String)"; }
} Else {


########################################################################################################################################################################################################
# Check for Added/Removed Software 
########################################################################################################################################################################################################
    #WriteLog -Log "Checking Software...";
    $EmailText = "";
    $SoftwareChange = 0;
    $OldSoftwareWithoutVersions = ($Record.SoftwareWithoutVersions).Split(";").Trim();
    $NewSoftwareWithoutVersions = ($DataObject.SoftwareWithoutVersions).Split(";").Trim();
    $Changes = Compare-Object -ReferenceObject $OldSoftwareWithoutVersions -DifferenceObject $NewSoftwareWithoutVersions;
    $RemovedSW = @{ 
        Name = 'Removed Software'; 
        Expression = { 
            $Index = [array]::indexof($OldSoftwareWithoutVersions,$_.InputObject);
            ($Record.SoftwareWithVersions).Split(";").Trim()[$Index];
        }
    }
    $AddedSW = @{ 
        Name = 'Added Software'; 
        Expression = { 
            $Index = [array]::indexof($NewSoftwareWithoutVersions,$_.InputObject);
            ($DataObject.SoftwareWithVersions).Split(";").Trim()[$Index];
        }
    }
    $RemovedSoftware = $Changes | Where-Object { $_.SideIndicator -eq '<='; } | Select-Object $RemovedSW;
    $AddedSoftware = $Changes | Where-Object { $_.SideIndicator -eq '=>'; } | Select-Object $AddedSW;
    If ($AddedSoftware) {
        $SoftwareChange = $SoftwareChange + 1;
        $EmailText += $AddedSoftware | Out-String;
    }
    If ($RemovedSoftware) {
        $SoftwareChange = $SoftwareChange + 2;
        $EmailText += $RemovedSoftware | Out-String;
    }
    If ($SoftwareChange -gt 0) {
        Switch ($SoftwareChange) {
            1 { $SoftwareChange = "Added"; }
            2 { $SoftwareChange = "Removed"; }
            3 { $SoftwareChange = "Added & Removed"; }
        }
        WriteLog -Log "[SOFTWARE] Software Change Found!" -Data $EmailText;
        EmailAlert -Subject "Software $($SoftwareChange)" -Body "$EmailText $($Changes | Out-String)";
    } Else { WriteLog -Log "[SOFTWARE] No Change in Software Found."; }
    

########################################################################################################################################################################################################
# Check for Major Changes Since Last Report 
########################################################################################################################################################################################################
    Try {
        $ToCompare = 'DeviceName','IpAddress','MacAddress','NetworkAdapters','CPU','RAM_Installed','Drives','DHCP','OS','Bios','LocalAdmins','RemoteUsers','Graphics','Webcam';
        $IpProp = @{ Name='IpAddress';  Expression={ ($_.IpAddress -replace '\[.*\]', '').Trim() } };
        $MacProp = @{ Name='MacAddress'; Expression={ ($_.MacAddress -replace '\[.*\]', '').Trim() } };
        $NetProp = @{ Name='NetworkAdapters'; Expression={ ($_.NetworkAdapters -replace '\[.*\]', '').Trim(); } };
        $ReferenceObject = $Record | Select-Object DeviceName,$IpProp,$MacProp,$NetProp,CPU,RAM_Installed,Drives,DHCP,OS,Bios,LocalAdmins,RemoteUsers,Graphics,Webcam;
        $DifferentObject = $DataObject | Select-Object DeviceName,$IpProp,$MacProp,$NetProp,CPU,RAM_Installed,Drives,DHCP,OS,Bios,LocalAdmins,RemoteUsers,Graphics,Webcam;
        $DataComparison = Compare-Object -ReferenceObject $ReferenceObject -DifferenceObject $DifferentObject -Property $ToCompare -IncludeEqual;
        $MajorChange = $DataComparison.SideIndicator;
        If ($MajorChange -ne "==") {
            $MajorChanges = "";
            Switch ($true) {
                ($Record.IpAddress -ne $DataObject.IpAddress) { $MajorChanges += "IP Address, "; }
                ($Record.MacAddress -ne $DataObject.MacAddress) { $MajorChanges += "MAC Address, ";}
                ($Record.NetworkAdapters -ne $DataObject.NetworkAdapters) { $MajorChanges += "Network Adapters, "; }
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
                WriteLog -Log "[CONFIGURATION] Major Configuration Change Found!";
                WriteLog -Log "[CONFIGURATION] Changes: $MajorChanges" -Data $($DataComparison | Format-List | Out-String);
                EmailAlert -Subject "$($MajorChanges) Changed" -Body "Change:  $MajorChanges`n$($DataComparison | Format-List | Out-String)";
            }
        } Else { WriteLog -Log "[CONFIGURATION] No Major Change in Device Configuration Found.";}
    } Catch { WriteLog -Log "[ERROR] Error Comparing Configurations." -Data $_; WriteLog -Data $DataComparison; }


########################################################################################################################################################################################################
# Save New Data to the Excel File
########################################################################################################################################################################################################
    $Record = $DataObject;
    Try {
        $Record | Export-Csv -Path $CsvFile;
        WriteLog -Log "[REPORT] $($Record.DeviceName) Self-Reported" -Data $Record;
    } Catch { WriteLog -Log "[ERROR] Error Saving Excel File." -Data $_; }
}


########################################################################################################################################################################################################
# SCCM Check-In
########################################################################################################################################################################################################
WriteLog -Log "Checking in to SCCM...";
If ($DataObject.OS -NotLike "*Server*") {
    Try {
        $SMSCli = [wmiclass] "root\ccm:sms_client";
        If (-NOT (Get-WmiObject -Namespace root\ccm -Class SMS_Client)) {
	        Stop-Service -Force winmgmt -ErrorAction SilentlyContinue;
   	        Set-Location  C:\Windows\System32\Wbem\;
   	        Remove-Item C:\Windows\System32\Wbem\Repository.old -Force -ErrorAction SilentlyContinue;
   	        Rename-Item Repository Repository.old -ErrorAction SilentlyContinue;
   	        Start-Service winmgmt;
        }
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
        Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000021}" -ErrorAction SilentlyContinue | Out-Null; 
        Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000121}" -ErrorAction SilentlyContinue | Out-Null; 
        Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000003}" -ErrorAction SilentlyContinue | Out-Null;
        Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000071}" -ErrorAction SilentlyContinue | Out-Null;
        Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000108}" -ErrorAction SilentlyContinue | Out-Null;
        Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000113}" -ErrorAction SilentlyContinue | Out-Null;
        Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000001}" -ErrorAction SilentlyContinue | Out-Null;
        If ($MachineStatus -AND $SoftwareStatus -AND $SoftwareDeployStatus) { 
            WriteLog -Log "[SUCCESS] SCCM Check-In Successful."; 
        } Else {
            $SMSCli.RepairClient();
            WriteLog -Log "[ERROR] SCCM Check-In Unsuccessful.";
        }
    } Catch { WriteLog -Log "[ERROR] Error Checking In with SCCM." -Data $_; }
}


########################################################################################################################################################################################################
# Remove all Local Group Policy and Refresh Domain Policy
########################################################################################################################################################################################################
#WriteLog -Log "Removing Local Group Policy...";
#If (Test-Path -Path "$env:windir\system32\GroupPolicyUsers") { Remove-Item "$env:windir\system32\GroupPolicyUsers" -Force -Recurse -ErrorAction SilentlyContinue; }
#If (Test-Path -Path "$env:windir\system32\GroupPolicy") { Remove-Item "$env:windir\system32\GroupPolicy" -Force -Recurse -ErrorAction SilentlyContinue; }
#gpupdate /force;


#########################################
# Th-th-th-th-that's all folks! 
#########################################
Write-Host "Script Completed in $([math]::Round((New-TimeSpan -Start $StartTime -End (Get-Date)).TotalSeconds, 1)) Seconds";
#exit 0;
#[Environment]::Exit(0);

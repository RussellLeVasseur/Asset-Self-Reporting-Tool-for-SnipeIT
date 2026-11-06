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
# Junk serial detection helper
########################################################################################################################################################################################################
function Test-InvalidInventorySerial {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $true
    }

    $v = $Value.Trim()

    if (
        $v -match '^To be filled by O\.?E\.?M\.?$' -or
        $v -match '^Default string$' -or
        $v -match '^System Serial Number$' -or
        $v -match '^Unknown$' -or
        $v -match '^None$'
    ) {
        return $true
    }

    return $false
}

########################################################################################################################################################################################################
# Static Variables 
########################################################################################################################################################################################################
If (!$ConfigFile) { 
    Write-Error "You did not provide a config file!";
    Write-Error "Use the flag '-ConfigFile' and provide a completed config."
    Write-Error "Exiting..."
    Start-Sleep -Seconds 10; Exit 1; 
}

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

# Script Version
$ScriptVersion = "1.10";

$StartTime = Get-Date;
$Today = Get-Date -UFormat "%d-%b-%Y";

$DeviceName = hostname;
[HashTable]$DataHashTable = @{};
$Win32_BIOS = Get-WmiObject -Class Win32_BIOS
$Win32_ComputerSystemProduct = Get-WmiObject -Class Win32_ComputerSystemProduct
$Win32_BaseBoard = Get-WmiObject -Class Win32_BaseBoard

$BiosSerial  = ($Win32_BIOS.SerialNumber | Out-String).Trim()
$UUID        = ($Win32_ComputerSystemProduct.UUID | Out-String).Trim()
$BoardSerial = ($Win32_BaseBoard.SerialNumber | Out-String).Trim()

# Optional override file (keep your existing logic)
$OverrideSerial = $null
if (Test-Path -Path 'C:\CCCJ\ASR\sn.txt' -PathType Leaf) {
    $OverrideSerial = (Get-Content -Path 'C:\CCCJ\ASR\sn.txt' | Select-Object -First 1).Trim()
}

if (-not (Test-InvalidInventorySerial $OverrideSerial)) {
    $SerialNumber = $OverrideSerial
}
elseif (-not (Test-InvalidInventorySerial $BiosSerial)) {
    $SerialNumber = $BiosSerial
}
elseif (
    -not [string]::IsNullOrWhiteSpace($UUID) -and
    $UUID -notmatch '^FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF$' -and
    $UUID -notmatch '^00000000-0000-0000-0000-000000000000$'
) {
    $SerialNumber = $UUID
}
elseif (-not (Test-InvalidInventorySerial $BoardSerial)) {
    $SerialNumber = "MB-$BoardSerial"
}
else {
    $SerialNumber = $DeviceName
}
$RandomNumber = Get-Random -Minimum 0 -Maximum 300;
$SystemInformation = Get-WmiObject -Namespace root\wmi -Class MS_SystemInformation;
$SoftwareLicensingServiceInfo = Get-WmiObject -query 'select * from SoftwareLicensingService';

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
#$DataHashTable.Add('SerialNumber', $SerialNumber);

$DataHashTable.Add('SerialNumber', $SerialNumber)

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

#Get UUID
$UUID = (Get-WmiObject -Class Win32_ComputerSystemProduct).UUID
$DataHashTable.Add('UUID', $UUID);

#Add Script Version
$DataHashTable.Add('ScriptVersion', $ScriptVersion);

#Get Bitlocker Status
$BitLockerRaw = Manage-BDE -Status | Out-String
$BitLockerVersionLine = $BitLockerRaw -split "`r?`n" | Where-Object {$_ -match 'BitLocker Version:'} | Select-Object -First 1
$BitLockerVersion = ($BitLockerVersionLine -replace 'BitLocker Version:\s*','') -as [string]
$DataHashTable.Add('BitLockerVersion', $BitLockerVersion);

$BitLockerLines = $BitLockerRaw -split "`r?`n" | Where-Object {
    $_ -match 'Conversion Status:' -or
    $_ -match 'Percentage Encrypted:' -or
    $_ -match 'Encryption Method:' -or
    $_ -match 'Protection Status:'
}
$BitLockerBDE = $BitLockerLines -join "`n"

$BitLockerVolRaw = Get-BitLockerVolume | ForEach-Object {
    $mount = $_.MountPoint
    $_.KeyProtector | ForEach-Object {
        "Drive: $mount | KeyProtectorId: $($_.KeyProtectorId) | KeyProtectorType: $($_.KeyProtectorType)"
    }
}
$BitLockerVolume = $BitLockerVolRaw -join "`n"
$BitLockerSummary = ($BitLockerBDE, $BitLockerVolume) -join "`n"
$DataHashTable.Add('BitLockerSummary', $BitLockerSummary);

#Get Secure Boot Status
Try {
    $SecureBootEnabled = Confirm-SecureBootUEFI -ErrorAction Stop
    If ($SecureBootEnabled -eq $true) {
        $SecureBootStatus = "Enabled"
    } ElseIf ($SecureBootEnabled -eq $false) {
        $SecureBootStatus = "Disabled"
    } Else {
        $SecureBootStatus = "Unknown"
    }
} Catch {
    $SecureBootStatus = "Not Supported (Legacy BIOS)"
}
$DataHashTable.Add('SecureBootStatus', $SecureBootStatus);

# --- Robust Secure Boot Certificate Expiry ---

Write-Host "Secure Boot Certificate Expiry" -ForegroundColor Green
$SecureBootCertInfo = ""

# GUIDs from UEFI spec
$GUID_X509   = [guid]'a5c059a1-94e4-4aa7-87b5-ab155c2bf072'  # EFI_CERT_X509_GUID
$GUID_SHA256 = [guid]'c1c41626-504c-4092-aca9-41f936934328'  # EFI_CERT_SHA256_GUID
$GUID_PKCS7  = [guid]'4aafd29d-68df-49ee-8aa9-347d375665a7'  # EFI_CERT_TYPE_PKCS7_GUID'

function Read-UInt32LE([byte[]]$b, [int]$o) { [BitConverter]::ToUInt32($b, $o) }
function Get-Guid([byte[]]$b, [int]$o) {
    $slice = New-Object byte[] 16
    [Buffer]::BlockCopy($b, $o, $slice, 0, 16)
    # Guid(byte[]) ctor expects little-endian fields as in UEFI storage
    return New-Object System.Guid(,$slice)
}

function Parse-Db {
    param([byte[]]$Raw)

    $result = New-Object System.Collections.Generic.List[object]
    $ofs = 0
    $HDR = 16 + 4 + 4 + 4  # Type(16) + ListSize + HeaderSize + SigSize

    while ($ofs -le $Raw.Length - $HDR) {
        $sigType = Get-Guid $Raw $ofs; $ofs += 16
        $listSize = Read-UInt32LE $Raw $ofs; $ofs += 4
        $hdrSize  = Read-UInt32LE $Raw $ofs; $ofs += 4
        $sigSize  = Read-UInt32LE $Raw $ofs; $ofs += 4

        # Bounds check
        $listStart = $ofs - $HDR
        $listEnd   = $listStart + $listSize
        if ($listSize -lt $HDR -or $listEnd -gt $Raw.Length -or $sigSize -lt 16) {
            # Malformed list; bail out of this list
            $ofs = $listEnd
            continue
        }

        # Skip signature header (rarely used)
        $ofs += $hdrSize

        # Walk EFI_SIGNATURE_DATA entries
        while ($ofs + $sigSize -le $listEnd) {
            # First 16 bytes is Owner GUID, remainder is SignatureData
            $owner = Get-Guid $Raw $ofs
            $ofsOwnerEnd = $ofs + 16
            $dataLen = $sigSize - 16

            $sigData = New-Object byte[] $dataLen
            [Buffer]::BlockCopy($Raw, $ofsOwnerEnd, $sigData, 0, $dataLen)
            $ofs += $sigSize

            switch ($true) {
                { $sigType -eq $GUID_X509 } {
                    try {
                        $x = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2(,$sigData)
                        $result.Add([pscustomobject]@{
                            Type        = 'X509'
                            OwnerGuid   = $owner
                            Subject     = $x.Subject
                            Issuer      = $x.Issuer
                            NotAfter    = $x.NotAfter
                            Thumbprint  = $x.Thumbprint
                        })
                    } catch {}
                }
                { $sigType -eq $GUID_PKCS7 } {
                    try {
                        Add-Type -AssemblyName System.Security # for Pkcs
                        Add-Type -AssemblyName System.Security.Cryptography
                        $cms = New-Object System.Security.Cryptography.Pkcs.SignedCms
                        $cms.Decode($sigData)
                        foreach ($c in $cms.Certificates) {
                            $result.Add([pscustomobject]@{
                                Type        = 'PKCS7-Cert'
                                OwnerGuid   = $owner
                                Subject     = $c.Subject
                                Issuer      = $c.Issuer
                                NotAfter    = $c.NotAfter
                                Thumbprint  = $c.Thumbprint
                            })
                        }
                    } catch {}
                }
                { $sigType -eq $GUID_SHA256 } {
                    # Hash allow-list entry; no expiry
                    $result.Add([pscustomobject]@{
                        Type        = 'SHA256-Hash'
                        OwnerGuid   = $owner
                        Subject     = '<hash entry>'
                        Issuer      = ''
                        NotAfter    = $null
                        Thumbprint  = ([BitConverter]::ToString($sigData) -replace '-', '')
                    })
                }
                default {
                    # Unknown/other signature type; record basics
                    $result.Add([pscustomobject]@{
                        Type        = "Other:$sigType"
                        OwnerGuid   = $owner
                        Subject     = ''
                        Issuer      = ''
                        NotAfter    = $null
                        Thumbprint  = ''
                    })
                }
            }
        }

        # move to next list (in case of padding/misalignment)
        $ofs = $listEnd
    }

    # De-duplicate by Thumbprint when present (PKCS7 can contain repeats)
    $seen = @{}
    $unique = foreach ($r in $result) {
        $key = if ($r.Thumbprint) { $r.Type + ':' + $r.Thumbprint } else { [guid]::NewGuid().ToString() }
        if (-not $seen.ContainsKey($key)) {
            $seen[$key] = $true
            $r
        }
    }
    return ,$unique
}

try {
    if ($SecureBootStatus -in @('Enabled','Disabled')) {
        $raw = (Get-SecureBootUEFI -Name db -ErrorAction Stop).Bytes
        Write-Host "  Parsing EFI signature lists ($($raw.Length) bytes)..." -ForegroundColor Cyan
        $entries = Parse-Db -Raw $raw

        $certs = $entries | Where-Object { $_.Type -in 'X509','PKCS7-Cert' }
        $hashes = $entries | Where-Object { $_.Type -eq 'SHA256-Hash' }

        Write-Host ("  Found {0} certificate(s), {1} hash entry/entries" -f $certs.Count, $hashes.Count) -ForegroundColor Cyan
        Write-Host ""

        if ($certs.Count -gt 0) {
            $today = Get-Date
            $lines = @()
            foreach ($c in ($certs | Sort-Object Subject, NotAfter -Unique)) {
                $days = if ($c.NotAfter) { ($c.NotAfter - $today).Days } else { $null }
                $expiry = if ($c.NotAfter) { $c.NotAfter.ToString('yyyy-MM-dd') } else { 'N/A' }
                $subjectCN = ($c.Subject -replace '^.*CN=([^,]+).*$','$1')

                if ($days -lt 0) { $status = "EXPIRED"; $color = "Red" }
                elseif ($days -lt 180) { $status = "EXPIRING SOON"; $color = "Yellow" }
                else { $status = "Valid"; $color = "Green" }

                Write-Host ("    {0}" -f $subjectCN) -ForegroundColor White
                Write-Host ("      Expires: {0} ({1}{2})" -f $expiry, $status, $(if ($days -ne $null) { ", $days days" } else { "" })) -ForegroundColor $color
                $lines += ("{0} : {1} ({2}{3})" -f $subjectCN, $expiry, $status, $(if ($days -ne $null) { ", $days days" } else { "" }))
            }

            # Your summary sink
            $SecureBootCertInfo = $lines -join "`n"

            # Your early-2026 warning:
            $exp26 = $certs | Where-Object { $_.NotAfter -and $_.NotAfter.Year -eq 2026 -and $_.NotAfter.Month -le 6 }
            if ($exp26.Count -gt 0) {
                Write-Host ""
                Write-Host ("  WARNING: Found {0} certificate(s) expiring in early 2026!" -f $exp26.Count) -ForegroundColor Red
            }
        }
        else {
            $SecureBootCertInfo = "No X.509 certificates present (DB may contain only hashes)."
            Write-Host "  $SecureBootCertInfo" -ForegroundColor Yellow
        }
    } else {
        $SecureBootCertInfo = "N/A (Secure Boot not supported)"
        Write-Host "  $SecureBootCertInfo" -ForegroundColor Gray
    }
} catch {
    $SecureBootCertInfo = "Unable to read/parse Secure Boot DB ($($_.Exception.Message))"
    Write-Host "  $SecureBootCertInfo" -ForegroundColor Yellow
}

# --- Pick a single meaningful value for Snipe-IT (text + date) ---

function Get-ConciseMicrosoftSecureBootExpiry {
    param(
        [Parameter(Mandatory)] [object[]] $Certs,            # entries with Subject / NotAfter
        [Parameter(Mandatory)] [string]   $SecureBootStatus  # "Enabled" | "Disabled" | etc.
    )

    if (-not $Certs -or $Certs.Count -eq 0) {
        return [pscustomobject]@{
            Text = 'No X.509 certificates (hash-only DB)'
            Date = $null
        }
    }

    # Microsoft CAs relevant to Windows boot trust
    $msCerts = $Certs | Where-Object {
        $_.Subject -match 'Microsoft' -and
        (
            $_.Subject -match 'Windows\s+Production\s+PCA' -or
            $_.Subject -match '\bUEFI\s+CA\b'              -or
            $_.Subject -match 'Option\s+ROM\s+UEFI\s+CA'
        )
    }

    if (-not $msCerts) {
        return [pscustomobject]@{
            Text = 'No Microsoft UEFI/Boot CAs found'
            Date = $null
        }
    }

    # Earliest expiry among MS CAs = practical “deadline”
    $earliest = $msCerts | Sort-Object NotAfter | Select-Object -First 1
    $today    = Get-Date
    $days     = if ($earliest.NotAfter) { ($earliest.NotAfter - $today).Days } else { $null }
    $expiry   = if ($earliest.NotAfter) { $earliest.NotAfter.ToString('yyyy-MM-dd') } else { 'N/A' }
    $subjectCN = ($earliest.Subject -replace '^.*CN=([^,]+).*$','$1')

    $status =
        if ($days -eq $null)   { 'N/A' }
        elseif ($days -lt 0)   { 'EXPIRED' }
        elseif ($days -lt 180) { 'EXPIRING SOON' }
        else                   { 'Valid' }

    $sbFlag = if ($SecureBootStatus -eq 'Enabled') { '' } else { '[SB OFF] ' }

    [pscustomobject]@{
        Text = ("{0}{1} : {2} ({3}{4})" -f $sbFlag, $subjectCN, $expiry, $status,
                $(if ($days -ne $null) { ", $days days" } else { "" }))
        Date = $earliest.NotAfter   # DateTime (or $null)
    }
}

# Produce BOTH fields the rest of the script will use
$sbSummary = Get-ConciseMicrosoftSecureBootExpiry -Certs $certs -SecureBootStatus $SecureBootStatus

# Long contextual text
$DataHashTable['SecureBootCertExpiry'] = $sbSummary.Text

# ISO date string (or empty if no date)
$DataHashTable['SecureBootCertExpiryDate'] = if ($sbSummary.Date) {
    $sbSummary.Date.ToString('yyyy-MM-dd')
} else {
    ''
}

Write-Host ("  Snipe-IT Secure Boot Cert Expiry (text): {0}" -f $DataHashTable['SecureBootCertExpiry']) -ForegroundColor Cyan
Write-Host ("  Snipe-IT Secure Boot Cert Expiry Date   : {0}" -f $DataHashTable['SecureBootCertExpiryDate']) -ForegroundColor Cyan

# --- Secure Boot Compliance classification ---

# Helper: read DB explicitly to test for 2023 CAs in Default DB (even if Active DB was parsed)
$entriesDefault = $null
try {
    $rawDefault = (Get-SecureBootUEFI -Name dbDefault -ErrorAction Stop).Bytes
    if ($rawDefault -and $rawDefault.Length -gt 0) {
        $entriesDefault = Parse-Db -Raw $rawDefault
    }
} catch {}

# Patterns for Microsoft CA subjects
$reMS2011 = 'Microsoft.*(UEFI\s*CA\s*2011|Windows\s+Production\s+PCA\s+2011)'
$reMS2023 = '(Windows\s+UEFI\s+CA\s+2023|Microsoft\s+UEFI\s+CA\s+2023|Microsoft\s+Corporation\s+KEK\s+2K\s+CA\s+2023|Microsoft\s+Option\s+ROM\s+UEFI\s+CA\s+2023)'

# Active DB (parsed to $certs earlier)
$ms2011Active  = @($certs | Where-Object { $_.Subject -match $reMS2011 })
$ms2023Active  = @($certs | Where-Object { $_.Subject -match $reMS2023 })

# Default DB (explicit check regardless of which source we parsed to $certs)
$ms2023Default = @($entriesDefault | Where-Object { $_.Subject -match $reMS2023 })

# Earliest 2011 expiry we can see (Active first; if none, consider Default)
$earliest2011 =
    @($ms2011Active + ($entriesDefault | Where-Object { $_.Subject -match $reMS2011 })) |
    Where-Object { $_.NotAfter } |
    Sort-Object NotAfter |
    Select-Object -First 1

$today = Get-Date
$daysTo2011 = if ($earliest2011 -and $earliest2011.NotAfter) { ($earliest2011.NotAfter - $today).Days } else { $null }

function Get-SBCompliance {
    param(
        [string] $SecureBootStatus,
        [array]  $Ms2023Active,
        [array]  $Ms2023Default,
        [Nullable[int]] $DaysTo2011
    )

    switch ($SecureBootStatus) {
        'Enabled' {
            if ($Ms2023Active.Count -gt 0) {
                return 'Compliant - 2023 CAs Active'
            }
            elseif ($Ms2023Default.Count -gt 0) {
                # Firmware will repopulate Active DB from Default DB on firmware/OS update or specific tasks
                return 'Pending - 2023 CA present in Default DB (Active DB still 2011)'
            }
            else {
                if ($DaysTo2011 -eq $null) {
                    return 'At Risk - 2023 CA missing; 2011 expiry unknown'
                }
                elseif ($DaysTo2011 -lt 0) {
                    return 'At Risk - 2011 CAs expired; 2023 CA missing'
                }
                elseif ($DaysTo2011 -lt 180) {
                    return ('At Risk - 2011 CAs expiring soon ({0} days); 2023 CA missing' -f $DaysTo2011)
                }
                else {
                    return ('At Risk - 2023 CA missing; 2011 CAs expire in {0} days' -f $DaysTo2011)
                }
            }
        }
        'Disabled' {
            return 'Secure Boot Disabled - Cert validity irrelevant'
        }
        'Not Supported (Legacy BIOS)' {
            return 'Not Supported (Legacy BIOS) - Compliance N/A'
        }
        default {
            # Includes 'Unknown' or cases where DB is unavailable (Setup Mode or keys cleared)
            return 'Unknown - DB not available (Setup Mode or keys not enrolled)'
        }
    }
}

$DataHashTable['SecureBootCompliance'] = Get-SBCompliance -SecureBootStatus $SecureBootStatus `
                                                         -Ms2023Active $ms2023Active `
                                                         -Ms2023Default $ms2023Default `
                                                         -DaysTo2011 $daysTo2011

Write-Host ("  Secure Boot Compliance: {0}" -f $DataHashTable['SecureBootCompliance']) -ForegroundColor Cyan

#Get BIOS Release Date
If ($Win32_BIOS.ReleaseDate) {
    Try {
        $BiosReleaseDate = [System.Management.ManagementDateTimeConverter]::ToDateTime($Win32_BIOS.ReleaseDate);
        $DataHashTable.Add('BiosReleaseDate', $BiosReleaseDate.ToString('yyyy-MM-dd'));
    } Catch {
        $DataHashTable.Add('BiosReleaseDate', '');
    }
} Else {
    $DataHashTable.Add('BiosReleaseDate', '');
}

$Domain = (Get-WmiObject -Class Win32_ComputerSystem).Domain;
$DataHashTable.Add('Domain', " $Domain");

#Get UI Language
$RegkeyResult = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Nls\Language" -Name "Default"

if ($RegkeyResult.Default -eq "0809") {
    $UILanguage = "en-GB"
} elseif ($RegkeyResult.Default -eq "0409") {
    $UILanguage = "en-US"
}
$DataHashTable.Add('WindowsUILanguage', $UILanguage);

#Get Monitor Information
$MonitorData = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorID

# "DEL DELL U2723QE [ABC123]" from a WmiMonitorID instance; also used by the Display Resolution section.
function Get-MonitorName {
    param([Object] $Monitor)
    $Manufacturer = ($Monitor.ManufacturerName | Where-Object {$_ -ne 0} | ForEach-Object {[char]$_}) -join ''
    $Model = ($Monitor.UserFriendlyName | Where-Object {$_ -ne 0} | ForEach-Object {[char]$_}) -join ''
    $Serial = ($Monitor.SerialNumberID | Where-Object {$_ -ne 0} | ForEach-Object {[char]$_}) -join ''
    "$Manufacturer $Model [$Serial]"
}

$MonitorInfo = ($MonitorData | ForEach-Object { Get-MonitorName -Monitor $_ }) -join "; "
$DataHashTable.Add('MonitorInfo', $MonitorInfo);
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
} ElseIf ($Win32_ComputerSystem.Manufacturer -like '*Dell*') {
    $RequiredModules = 'SnipeitPS', 'DellBIOSProvider', 'ActiveDirectory', 'PSWindowsUpdate';
} Else {
    $RequiredModules = 'SnipeitPS', 'ActiveDirectory', 'PSWindowsUpdate';
}
# Detect OS type once to decide how to install RSAT/AD tools
$osInfo = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
$isServer = $false
If ($osInfo) { $isServer = ($osInfo.ProductType -ne 1) }

$RequiredModules | ForEach-Object {
    Try {
        $Mdle = $_
        If (!(Get-Module -ListAvailable -Name $Mdle)) {
            WriteLog -Log "$Mdle not found. Installing..."
            If ($Mdle -eq 'ActiveDirectory') {
                If ($isServer) {
                    If (Get-Command -Name Install-WindowsFeature -ErrorAction SilentlyContinue) {
                        Try {
                            Install-WindowsFeature RSAT-AD-PowerShell -ErrorAction Stop
                            WriteLog -Log "Installed RSAT-AD-PowerShell via Install-WindowsFeature."
                        } Catch {
                            WriteLog -Log "[ERROR] Failed to install RSAT via Install-WindowsFeature. $_"
                        }
                    } Else {
                        WriteLog -Log "[WARN] Install-WindowsFeature not available on this Server. Skipping RSAT AD install."
                    }
                } Else {
                    If (Get-Command -Name Add-WindowsCapability -ErrorAction SilentlyContinue) {
                        Try {
                            Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0" -ErrorAction Stop
                            WriteLog -Log "Installed RSAT ActiveDirectory via Add-WindowsCapability."
                        } Catch {
                            WriteLog -Log "[ERROR] Failed to install RSAT via Add-WindowsCapability. $_"
                        }
                    } Else {
                        WriteLog -Log "[WARN] Add-WindowsCapability not available on this Client. Skipping RSAT AD install."
                    }
                }
            } Else {
                Try {
                    Install-Module -Name $Mdle -Force -AllowClobber -ErrorAction Stop
                    WriteLog -Log "Installed PowerShell module $Mdle."
                } Catch {
                    WriteLog -Log "[ERROR] Failed to Install-Module $Mdle. $_"
                    EmailAlert -Subject "[ERROR] Installing Module" -Body "$( $_ | Out-String)"
                }
            }
        } Else {
            Try {
                $Latest = [String](Find-Module -Name $Mdle -ErrorAction SilentlyContinue | Sort-Object Version -Descending | Select-Object -First 1).version
                $Installed = [String](Get-Module -ListAvailable -Name $Mdle | Select-Object -First 1).version
                If ($Latest -and [System.Version]$Latest -gt [System.Version]$Installed) {
                    WriteLog -Log "[UPDATE] Updating $($Mdle)..."
                    Update-Module -Name $Mdle -Force
                }
            } Catch { }
        }

        Try { Import-Module -Name $Mdle -Force } Catch {
            WriteLog -Log "[ERROR] Unable to Import $($Mdle) Module." -Data $_
            EmailAlert -Subject "[ERROR] Importing Module" -Body "$( $_ | Out-String)"
        }
    } Catch { WriteLog -Log "[ERROR] $($_ | Out-String)" }
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
$DataHashTable.Add('SKU', $SystemInformation.SystemSKU);

########################################################################################################################################################################################################
# Operating System Information
########################################################################################################################################################################################################
#WriteLog -Log "Gathering Operating System Information...";
$Win32_OperatingSystem = Get-WmiObject -Class Win32_OperatingSystem;
$DataHashTable.Add('OS', ($Win32_OperatingSystem.Name).Split("|")[0]);
$DataHashTable.Add('Build', $Win32_OperatingSystem.Version);
$Win_Version = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").DisplayVersion
$DataHashTable.Add('WindowsVersion', $Win_Version);
$DataHashTable.Add('BIOSWindowsLicenseKey', $SoftwareLicensingServiceInfo.OA3xOriginalProductKey);
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
    $MacAddress += "$($_.MacAddress -replace '-',':')";
    If ($_.Status -eq 'Up') {
        $InterfaceAlias = "$($_.Name)";
        $IpAddress += "$((Get-NetIpAddress | Where-Object { $_.AddressFamily -Like "IPv4" -and $_.InterfaceAlias -eq $InterfaceAlias; }).IPAddress)";
    }
}
$MacAddress = $MacAddress -join "`n";
$IpAddress = $IpAddress -join "`n";
$NetworkAdapters = $NetworkAdapters -join "`n";
$DataHashTable.Add('IpAddress', $IpAddress);
$DataHashTable.Add('MacAddress', $MacAddress);
$DataHashTable.Add('NetworkAdapters', $NetworkAdapters);

# Custom MAC Address-based device identification
$BlackviewMacAddresses = @("68:1D:EF:50:47:02", "8C:EA:12:98:5D:AC");
$CollectedMacAddresses = $MacAddress -split "`n";
$MacMatch = $false;
ForEach ($CollectedMac in $CollectedMacAddresses) {
    If ($BlackviewMacAddresses -contains $CollectedMac) {
        $MacMatch = $true;
        Break;
    }
}
If ($MacMatch) {
    WriteLog -Log "[LOG] Blackview MP60 MAC Address Detected. Overriding Model, Manufacturer and Serial Number.";
    $DataHashTable['Model'] = "Blackview MP60";
    $DataHashTable['Manufacturer'] = "Blackview";
    $DataHashTable['SerialNumber'] = $DeviceName;  # Use hostname as serial number for Blackview devices
}
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
# $LastUser = Get-WmiObject Win32_NetworkLoginProfile | Where{$_.LastLogon} | Sort LastLogon -Descending | Select-Object Name -first 1;
# Doesn't seem to work # $LastUser = (Get-CimInstance -ClassName Win32_ComputerSystem).Username
# Doesn't report anything when running as SYSTEM # $LastUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$LastUser = Get-ChildItem "c:\Users" | Sort-Object LastWriteTime -Descending | Select-Object Name, LastWriteTime -first 1

$DataHashTable.Add('LastUser', ($LastUser).Name);

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
# Display Resolution, Scaling and Usable Space
########################################################################################################################################################################################################
# One line per active display, e.g.
#   DEL DELL U2723QE [ABC123]: 1920x1080 @100% = 1920x1080 usable (native 3840x2160) primary LOW-RES
# LOW-RES     = desktop resolution is below the panel's native mode.
# UNDER-1080P = usable space (resolution divided by Windows scaling) is below 1920x1080.
# The script runs as SYSTEM, which cannot see the user's desktop, so:
#   - native mode comes from each monitor's EDID (WmiMonitorListedSupportedSourceModes)
#   - current mode comes from the newest layout Windows saved under GraphicsDrivers\Configuration
#   - scaling comes from the signed-in user's AppliedDPI, which Windows records for the primary display at sign-in;
#     other displays show "@?" (their stored scaling is an offset from an undocumented "recommended" value), and so
#     does the primary if the layout changed after sign-in.
# The registry layout details are community-documented, not Microsoft-documented; the script log records which layout was used.
#WriteLog -Log "Gathering Display Resolution Information...";
$MinUsableWidth = 1920;
$MinUsableHeight = 1080;

# "DISPLAY\DEL40F5\5&2b1b1c0b&0&UID4357_0" -> "DEL40F5"
function Get-PnpCodeFromInstanceName {
    param([String] $InstanceName)
    $Parts = $InstanceName -split '\\';
    If ($Parts.Count -ge 2) { return $Parts[1].ToUpper(); }
    return '';
}

# The monitor's preferred (native) mode from its EDID, or $null if it doesn't report one.
function Get-NativeDisplayMode {
    param([Object] $SourceModes)
    If (-NOT $SourceModes -OR -NOT $SourceModes.MonitorSourceModes) { return $null; }
    $Modes = @($SourceModes.MonitorSourceModes);
    $Index = [int]$SourceModes.PreferredMonitorSourceModeIndex;
    If ($Index -lt 0 -OR $Index -ge $Modes.Count) { return $null; }
    $Mode = $Modes[$Index];
    If (-NOT $Mode.HorizontalActivePixels -OR -NOT $Mode.VerticalActivePixels) { return $null; }
    return [PSCustomObject]@{ Width = [int]$Mode.HorizontalActivePixels; Height = [int]$Mode.VerticalActivePixels };
}

# A layout's SetId names one EDID-based ID per display, each starting with the 3-letter manufacturer and 4-hex
# product code that WmiMonitorID's InstanceName also uses. Returns those codes in order.
function Get-PnpCodesFromSetId {
    param([String] $SetId)
    return @([regex]::Matches($SetId, '(?i)(?<![A-Z0-9])[A-Z]{3}[0-9A-F]{4}') | ForEach-Object { $_.Value.ToUpper() });
}

# AppliedDPI 96 = 100%, 120 = 125%, 144 = 150% ...
function ConvertTo-ScalePercent {
    param([Object] $Dpi)
    If ($null -eq $Dpi -OR [int]$Dpi -le 0) { return $null; }
    return [int][math]::Round([int]$Dpi * 100 / 96);
}

function Format-DisplayLine {
    param(
        [String] $Name, [int] $Width, [int] $Height, [Object] $ScalePercent, [Object] $Native, [Bool] $IsPrimary,
        [int] $MinWidth = 1920, [int] $MinHeight = 1080
    )
    $Line = "$($Name): $($Width)x$($Height)";
    $Flags = @();
    If ($ScalePercent) {
        $UsableWidth = [int][math]::Floor($Width * 100 / $ScalePercent);
        $UsableHeight = [int][math]::Floor($Height * 100 / $ScalePercent);
        $Line += " @$($ScalePercent)% = $($UsableWidth)x$($UsableHeight) usable";
    } Else {
        # Scaling unknown: the resolution itself is the most usable space there can be.
        $UsableWidth = $Width;
        $UsableHeight = $Height;
        $Line += " @?";
    }
    If ($Native) {
        # Compare long side with long side so a portrait-rotated panel isn't flagged for being rotated.
        $CurrentLong = [math]::Max($Width, $Height); $CurrentShort = [math]::Min($Width, $Height);
        $NativeLong = [math]::Max($Native.Width, $Native.Height); $NativeShort = [math]::Min($Native.Width, $Native.Height);
        If ($CurrentLong -lt $NativeLong -OR $CurrentShort -lt $NativeShort) {
            $Line += " (native $($Native.Width)x$($Native.Height))";
            $Flags += 'LOW-RES';
        } Else {
            $Line += " (native)";
        }
    } Else {
        $Line += " (native ?)";
    }
    If ($IsPrimary) { $Line += " primary"; }
    If ($UsableWidth -lt $MinWidth -OR $UsableHeight -lt $MinHeight) { $Flags += 'UNDER-1080P'; }
    If ($Flags.Count -gt 0) { $Line += " $($Flags -join ' ')"; }
    return $Line;
}

# $Displays: the newest layout's displays in order (Width, Height, X, Y).
# $Monitors: PnpCode, Name, Native for each connected monitor.
# Displays are matched to monitors through the layout's SetId; if that can't be done, lines are named "Display N".
function Format-DisplayReport {
    param(
        [Object[]] $Displays, [Object[]] $Monitors, [String] $SetId, [Object] $PrimaryScalePercent,
        [int] $MinWidth = 1920, [int] $MinHeight = 1080
    )
    # @() so a single-display SetId stays an array (otherwise $Codes[0] is its first character).
    $Codes = @(Get-PnpCodesFromSetId -SetId $SetId);
    $CanMatch = ($Codes.Count -eq $Displays.Count);
    $Taken = @{};
    $Lines = for ($i = 0; $i -lt $Displays.Count; $i++) {
        $Display = $Displays[$i];
        $Monitor = $null;
        If ($CanMatch) {
            for ($m = 0; $m -lt $Monitors.Count; $m++) {
                If (-NOT $Taken.ContainsKey($m) -AND $Monitors[$m].PnpCode -eq $Codes[$i]) { $Monitor = $Monitors[$m]; $Taken[$m] = $true; break; }
            }
        }
        $Name = If ($Monitor -AND "$($Monitor.Name)" -match '[A-Za-z0-9]') { "$($Monitor.Name)".Trim() } Else { "Display $($i + 1)" };
        $Native = If ($Monitor) { $Monitor.Native } Else { $null };
        $IsPrimary = ($Display.X -eq 0 -AND $Display.Y -eq 0);
        $Scale = If ($IsPrimary) { $PrimaryScalePercent } Else { $null };
        Format-DisplayLine -Name $Name -Width $Display.Width -Height $Display.Height -ScalePercent $Scale -Native $Native -IsPrimary $IsPrimary -MinWidth $MinWidth -MinHeight $MinHeight;
    }
    return (@($Lines) -join "`n");
}

# The layout in use now: the newest one with a timestamp and displays. Layouts with no real monitor (NOEDID_, SIMULATED_),
# e.g. saved while monitors were asleep, are only used if there is nothing else.
function Select-CurrentDisplayLayout {
    param([Object[]] $Layouts)
    $MaxFileTime = [uint64][DateTime]::MaxValue.ToFileTimeUtc();
    $Usable = @($Layouts | Where-Object { $_.Timestamp -gt 0 -AND $_.Timestamp -le $MaxFileTime -AND @($_.Displays).Count -gt 0 });
    $Real = @($Usable | Where-Object { $_.SetId -notmatch 'NOEDID|SIMULATED' });
    $Candidates = If ($Real.Count -gt 0) { $Real } Else { $Usable };
    return ($Candidates | Sort-Object -Property Timestamp -Descending | Select-Object -First 1);
}

# AppliedDPI is the scaling at sign-in. If the display layout changed after sign-in (e.g. the user docked), it may
# belong to a different primary display, so report it only when the layout is older than the sign-in.
function Get-PrimaryScaleForLayout {
    param([Object] $ScalePercent, [Object] $SignedInUtc, [uint64] $LayoutTimestamp)
    If (-NOT $ScalePercent -OR -NOT $SignedInUtc -OR $LayoutTimestamp -eq 0) { return $null; }
    If ([DateTime]::FromFileTimeUtc([int64]$LayoutTimestamp) -gt $SignedInUtc) { return $null; }
    return $ScalePercent;
}

# Every saved layout under GraphicsDrivers\Configuration: its SetId, Timestamp (FILETIME) and displays in order.
function Read-DisplayLayouts {
    $Root = 'HKLM:\SYSTEM\CurrentControlSet\Control\GraphicsDrivers\Configuration';
    Get-ChildItem -Path $Root -ErrorAction Stop | ForEach-Object {
        $LayoutKey = $_;
        $Layout = Get-ItemProperty -Path $LayoutKey.PSPath -ErrorAction Stop;
        $Stamp = [uint64]0;
        If ($Layout.Timestamp -is [byte[]] -AND $Layout.Timestamp.Length -ge 8) { $Stamp = [BitConverter]::ToUInt64($Layout.Timestamp, 0); }
        ElseIf ($null -ne $Layout.Timestamp) { $Stamp = [uint64]$Layout.Timestamp; }
        $SetId = If ($Layout.SetId) { $Layout.SetId } Else { ($LayoutKey.PSChildName -split '\^')[0] };
        # Each numbered subkey is a desktop surface; each of its numbered subkeys is a display showing it.
        $LayoutDisplays = @(Get-ChildItem -Path $LayoutKey.PSPath -ErrorAction Stop | Sort-Object PSChildName | ForEach-Object {
            $Surface = Get-ItemProperty -Path $_.PSPath -ErrorAction Stop;
            Get-ChildItem -Path $_.PSPath -ErrorAction Stop | Sort-Object PSChildName | ForEach-Object {
                $Target = Get-ItemProperty -Path $_.PSPath -ErrorAction Stop;
                $Width = If ($Target.'PrimSurfSize.cx') { $Target.'PrimSurfSize.cx' } Else { $Surface.'PrimSurfSize.cx' };
                $Height = If ($Target.'PrimSurfSize.cy') { $Target.'PrimSurfSize.cy' } Else { $Surface.'PrimSurfSize.cy' };
                If ($Width -AND $Height) {
                    [PSCustomObject]@{ Width = [int]$Width; Height = [int]$Height; X = [int]$Surface.'Position.cx'; Y = [int]$Surface.'Position.cy' };
                }
            }
        });
        [PSCustomObject]@{ SetId = $SetId; Timestamp = $Stamp; Displays = $LayoutDisplays };
    }
}

# The single signed-in user's sign-in scaling (ScalePercent) and sign-in time (SignedInUtc, from their first explorer.exe),
# or $null if nobody, or more than one user, is signed in. If explorer.exe was restarted, SignedInUtc is the restart time.
function Get-SignedInUserScale {
    $Shells = @(Get-CimInstance -ClassName Win32_Process -Filter "Name='explorer.exe'" -ErrorAction Stop | ForEach-Object {
        [PSCustomObject]@{ Sid = (Invoke-CimMethod -InputObject $_ -MethodName GetOwnerSid -ErrorAction Stop).Sid; Started = $_.CreationDate }
    } | Where-Object { $_.Sid });
    $Sids = @($Shells.Sid | Sort-Object -Unique);
    If ($Sids.Count -ne 1) {
        If ($Sids.Count -gt 1) { WriteLog -Log "[Displays] $($Sids.Count) users signed in; scaling not reported."; }
        return $null;
    }
    $WindowMetrics = "Registry::HKEY_USERS\$($Sids[0])\Control Panel\Desktop\WindowMetrics";
    If (-NOT (Test-Path -Path $WindowMetrics)) { WriteLog -Log "[Displays] No WindowMetrics key for signed-in user $($Sids[0]); scaling not reported."; return $null; }
    $SignedIn = ($Shells | Sort-Object -Property Started | Select-Object -First 1).Started;
    return [PSCustomObject]@{
        ScalePercent = ConvertTo-ScalePercent -Dpi (Get-ItemProperty -Path $WindowMetrics -ErrorAction Stop).AppliedDPI;
        SignedInUtc  = If ($SignedIn) { ([DateTime]$SignedIn).ToUniversalTime() } Else { $null };
    };
}

$DisplayResolution = 'Unknown';
Try {
    $SourceModes = @();
    If (@($MonitorData).Count -gt 0) {
        Try {
            $SourceModes = @(Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorListedSupportedSourceModes -ErrorAction Stop);
        } Catch { WriteLog -Log "[ERROR] [Displays] Could not read monitor native modes." -Data $_; }
    }
    $Monitors = @($MonitorData | ForEach-Object {
        $Instance = $_.InstanceName;
        [PSCustomObject]@{
            PnpCode = Get-PnpCodeFromInstanceName -InstanceName $Instance;
            Name    = Get-MonitorName -Monitor $_;
            Native  = Get-NativeDisplayMode -SourceModes ($SourceModes | Where-Object { $_.InstanceName -eq $Instance } | Select-Object -First 1);
        }
    });
    $CurrentLayout = Select-CurrentDisplayLayout -Layouts @(Read-DisplayLayouts);
    If ($CurrentLayout) {
        WriteLog -Log "[Displays] Using layout $($CurrentLayout.SetId) saved $([DateTime]::FromFileTimeUtc([int64]$CurrentLayout.Timestamp).ToString('u'))";
        $PrimaryScale = $null;
        Try {
            $UserScale = Get-SignedInUserScale;
            If ($UserScale) {
                If (-NOT $UserScale.ScalePercent) { WriteLog -Log "[Displays] Signed-in user has no AppliedDPI; scaling not reported."; }
                ElseIf (-NOT $UserScale.SignedInUtc) { WriteLog -Log "[Displays] Sign-in time unknown; scaling not reported."; }
                Else {
                    $PrimaryScale = Get-PrimaryScaleForLayout -ScalePercent $UserScale.ScalePercent -SignedInUtc $UserScale.SignedInUtc -LayoutTimestamp $CurrentLayout.Timestamp;
                    If (-NOT $PrimaryScale) { WriteLog -Log "[Displays] Display layout changed after sign-in; scaling not reported."; }
                }
            }
        } Catch { WriteLog -Log "[ERROR] [Displays] Could not read the signed-in user's scaling." -Data $_; }
        $DisplayResolution = Format-DisplayReport -Displays $CurrentLayout.Displays -Monitors $Monitors -SetId $CurrentLayout.SetId -PrimaryScalePercent $PrimaryScale -MinWidth $MinUsableWidth -MinHeight $MinUsableHeight;
        If ($DisplayResolution -match '(?m)^Display \d+:') {
            WriteLog -Log "[Displays] Some displays could not be matched to a monitor (layout $($CurrentLayout.SetId); monitors: $(@($Monitors.PnpCode) -join ', ')).";
        }
    } Else {
        WriteLog -Log "[Displays] No usable saved display layout found.";
    }
} Catch { WriteLog -Log "[ERROR] [Displays] Could not build the display resolution report." -Data $_; }
$DataHashTable.Add('DisplayResolution', $DisplayResolution);


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
If ($UseDell -eq $true) {

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
}

########################################################################################################################################################################################################
# Update SnipeIT 
########################################################################################################################################################################################################
WriteLog -Log "Checking in to SnipeIT...";
Start-Sleep -Seconds $RandomNumber;
Connect-SnipeitPS -URL $Snipe.Url -apiKey $Snipe.Token;
$SnipeAsset = Get-SnipeItAsset -asset_serial $DataHashTable['SerialNumber'];
WriteLog -Log "[DEBUG] Asset lookup result for serial '$($DataHashTable['SerialNumber'])': Count=$($SnipeAsset.Count), StatusCode=$($SnipeAsset.StatusCode)";
If ($SnipeAsset.StatusCode -eq 'InternalServerError') {
    WriteLog -Log "[WARN] First asset lookup failed with InternalServerError, retrying...";
    $SnipeAsset = Get-SnipeItAsset -asset_serial $DataHashTable['SerialNumber'];
    If ($SnipeAsset.StatusCode -eq 'InternalServerError') {
        EmailAlert -Subject "Error searching SnipeIT" -Body "(Duplicate Check)`n$($SnipeAsset | Format-List | Out-String)";
        Exit 0;
    }
}
## These custom fields have to be created and then allocated to a fieldset that is applied to the model.
## Create the custom field within SnipeIT and then ensure the field name is correct in the script below.
## The number on the end of the field name is simply the unique ID for the field, created at creation time
Switch ($true) {
    (![string]::IsNullOrEmpty($DataHashTable['Purchased'])) {
        $CustomValues.Add('purchase_date', $DataHashTable['Purchased'])
        Break
    }
    default { }  # Do nothing if 'Purchased' is null or empty
}
$CustomValues.Add('warranty_months', $DataHashTable['WarrantyMonths']);
# Use only the first MAC address for SnipeIT field (single MAC address validation)
$PrimaryMacAddress = ($DataHashTable['MacAddress'] -split "`n" | Where-Object { $_ -ne '' })[0];
If ($PrimaryMacAddress -and $PrimaryMacAddress -is [string]) {
    # Ensure it's a properly formatted single string value
    $CleanMacAddress = $PrimaryMacAddress.Trim() -as [string]
    If ($CleanMacAddress -match '^([0-9A-Fa-f]{2}[:-]){5}([0-9A-Fa-f]{2})$') {
        $CustomValues.Add('_snipeit_mac_address_1', $CleanMacAddress);
    }
}
$CustomValues.Add('_snipeit_cpu_2', $DataHashTable['CPU']);
$CustomValues.Add('_snipeit_ram_3', $DataHashTable['RAM']);
# Safe DNS lookup using first IP address
Try {
    $FirstIpAddress = ($DataHashTable['IpAddress'] -split "`n" | Where-Object { $_ -ne '' })[0];
    If ($FirstIpAddress) {
        $FQDN = [System.Net.Dns]::GetHostEntry($FirstIpAddress).HostName;
    } Else {
        $FQDN = $DataHashTable['DeviceName'];
    }
} Catch {
    $FQDN = $DataHashTable['DeviceName'];
}
$CustomValues.Add('_snipeit_fqdn_4', $FQDN);
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
$CustomValues.Add('_snipeit_ram_installed_20', $DataHashTable['RAM_Installed']);
$CustomValues.Add('_snipeit_webcam_22', $DataHashTable['Webcam']);
$CustomValues.Add('_snipeit_applied_updates_23', $DataHashTable['AppliedUpdates']);
$CustomValues.Add('_snipeit_network_adapters_24', $DataHashTable['NetworkAdapters']);
$CustomValues.Add('_snipeit_age_25', $DataHashTable['Age']);
$CustomValues.Add('_snipeit_last_logged_in_user_26', $DataHashTable['LastUser']);
$CustomValues.Add('_snipeit_operating_system_build_6', $DataHashTable['Build']);
$CustomValues.Add('_snipeit_windows_version_10', $DataHashTable['WindowsVersion']);
$CustomValues.Add('_snipeit_sku_7', $DataHashTable['SKU']);
$CustomValues.Add('_snipeit_uuid_35', $DataHashTable['UUID']);
# Ensure BitLocker version is a single string, not an array
$BitLockerVersionValue = $DataHashTable['BitLockerVersion'] -as [string]
$CustomValues.Add('_snipeit_bitlocker_version_36', $BitLockerVersionValue);
$CustomValues.Add('_snipeit_bitlocker_summary_38', $DataHashTable['BitLockerSummary']);
$CustomValues.Add('_snipeit_domain_39', $DataHashTable['Domain']);
$CustomValues.Add('_snipeit_windows_ui_language_40', $DataHashTable['WindowsUILanguage']);
$CustomValues.Add('_snipeit_bios_windows_license_key_8', $DataHashTable['BIOSWindowsLicenseKey']);
$CustomValues.Add('_snipeit_secure_boot_status_87', $DataHashTable['SecureBootStatus']);
$CustomValues.Add('_snipeit_secure_boot_cert_expiry_88', $DataHashTable['SecureBootCertExpiry']);
$CustomValues.Add('_snipeit_bios_release_date_89', $DataHashTable['BiosReleaseDate']);
$CustomValues.Add('_snipeit_script_version_90', $DataHashTable['ScriptVersion']);
# Add the new date field "Secure Boot Cert Expiry Date" as a date field
$CustomValues.Add('_snipeit_secure_boot_cert_expiry_date_91', $DataHashTable['SecureBootCertExpiryDate']);
# TEXT field in Snipe-IT: "Secure Boot Compliance"
$CustomValues.Add('_snipeit_secure_boot_compliance_92', $DataHashTable['SecureBootCompliance']);
$CustomValues.Add('_snipeit_monitors_93', $DataHashTable['MonitorInfo']);
$CustomValues.Add('_snipeit_display_resolution_94', $DataHashTable['DisplayResolution']);


$NextAuditDate = Get-Date;
If ($NextAuditDate.Month -ne 1) {
    $NextAuditDate = New-Object DateTime(($NextAuditDate.Year+1), 1, [DateTime]::DaysInMonth($NextAuditDate.Year, $NextAuditDate.Month))
    $Diff = ([int] [DayOfWeek]::Friday) - ([int]$lastDay.DayOfWeek);
    $NextAuditDate = @((Get-Date -Date $NextAuditDate.AddDays($Diff)),(Get-Date -Date $NextAuditDate.AddDays(- (7-$Diff))))[($Diff -ge 0)];
    $CustomValues.Add('next_audit_date', ($NextAuditDate | Get-Date -UFormat "%Y-%m-%d"));
}

If (!$SnipeAsset) {
    ## If no asset exists, go hunting for data and create one
    Try {
        Try {
            WriteLog -Log "[SnipeIT] Checking for manufacturer " + $($DataHashTable['Manufacturer']);
            $Manufacturer = Get-SnipeItManufacturer -search $DataHashTable['Manufacturer'];
            If (!$Manufacturer) { $Manufacturer = New-SnipeItManufacturer -name $DataHashTable['Manufacturer']; }
            $ManufacturerID = $Manufacturer.id;
        } Catch { WriteLog -Log "[SnipeIT] [ERROR] Unable to obtain Manufacturer ID." -Data $_; }
        Try {
            $String = "[SnipeIT] Checking for model " + $($DataHashTable['Model']);
            WriteLog -Log $String;
            $Model = Get-SnipeItModel -all | Where-Object { $_.name -eq "$($DataHashTable['Model'])" };
            $ModelData = $Model.notes -replace "&quot;",'"'
            # | ConvertFrom-Json;
            $String = "[SnipeIT] Retrieved " + $Model + " from SnipeIT";
            WriteLog -Log $String;
            If ($ModelData.LatestBios -gt $DataHashTable['Bios']) {
                ##########################################
            }
            If ($Model.count -eq 0) {
                If ($DataHashTable['OS'] -Contains "Server") { $ModelCatID = $Snipe.ServerCatID; } 
                Else { $ModelCatID = $Snipe.WorkstationCatID; }
                WriteLog -Log "[SnipeIT] Creating a new model in Snipe-IT";
                $Model = New-SnipeItModel -name $DataHashTable['Model'] -manufacturer_id $ManufacturerID -fieldset_id $Snipe.FieldSetID -category_id $ModelCatID;
            }
        } Catch { WriteLog -Log "[SnipeIT] [ERROR] Unable to obtain Model ID." -Data $_; }
        $SnipeAsset = New-SnipeItAsset -name $DataHashTable['DeviceName'] -status_id $Snipe.DefStatusID -model_id $Model.id -serial $DataHashTable['SerialNumber'] -asset_tag $DataHashTable['SerialNumber'] -customfields $CustomValues;
        If ($SnipeAsset -and $SnipeAsset.id) {
            WriteLog -Log "[SnipeIT] Created a new Asset in SnipeIT with ID: $($SnipeAsset.id)";
        } Else {
            WriteLog -Log "[SnipeIT] [ERROR] Asset creation returned invalid result.";
            WriteLog -Log "[DEBUG] Asset creation result:" -Data $SnipeAsset;
        }
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
        If ($SnipeAsset -and $SnipeAsset.id) {
            $UpdatedAsset = Set-SnipeItAsset -name $DataHashTable['DeviceName'] -id $SnipeAsset.id -status_id $Snipe.DefStatusID -customfields $CustomValues;
            WriteLog -Log "[SnipeIT] Updated an Asset in SnipeIT." -Data $UpdatedAsset;
            $SnipeAsset = $UpdatedAsset;
        } Else {
            WriteLog -Log "[ERROR] Cannot update asset - SnipeAsset or SnipeAsset.id is null/empty.";
            WriteLog -Log "[DEBUG] SnipeAsset content:" -Data $SnipeAsset;
        }
    } Catch { WriteLog -Log "[ERROR] Unable to Update SnipeIT Asset." -Data $_; }
}

# Mark asset as audited if machine has an IP address in the 10.2.0.0/16 range
$AuditSubnetIp = $DataHashTable['IpAddress'] -split "`n" | Where-Object { $_ -match '^10\.2\.' } | Select-Object -First 1;
If ($AuditSubnetIp -and $SnipeAsset -and $SnipeAsset.asset_tag) {
    Try {
        New-SnipeitAudit -tag $SnipeAsset.asset_tag;
        WriteLog -Log "[SnipeIT] Marked asset as audited (IP $AuditSubnetIp is in 10.2.0.0/16 range).";
    } Catch { WriteLog -Log "[ERROR] Unable to mark asset as audited in SnipeIT." -Data $_; }
}

# Check for Duplicate objects in SnipeIT
$DuplicateNames = Get-SnipeItAsset -Search $DataHashTable['DeviceName'] | Where-Object { $_.serial -ne $DataHashTable['SerialNumber'] -AND $_.assigned_to.name -NotLike "$($DeviceName)*" };
If ($DuplicateNames.StatusCode -eq 'InternalServerError') {
    EmailAlert -Subject "Error searching SnipeIT" -Body "(Duplicate Check)`n$($DuplicateNames | Format-List | Out-String)";
} ElseIf ($DuplicateNames.Count -gt 0) {
    ForEach ($Duplicate in $DuplicateNames) {
        $UpdatedAsset = Set-SnipeItAsset -name '' -id $Duplicate.id -customfields @{ _snipeit_ip_address_9 = ''; };
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
If ($UseSCCM -eq $true) {

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

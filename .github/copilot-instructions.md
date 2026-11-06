# GitHub Copilot Instructions for Asset-Self-Reporting-SnipeIT

## Repository Overview

This is a **Windows PowerShell automation script** designed to collect asset information from Windows machines and automatically update a SnipeIT inventory system. The script runs on individual Windows computers (workstations and servers) to self-report their hardware and software configurations.

**Key Details:**
- **Language:** PowerShell (UTF-8 encoded)
- **Primary File:** `AssetSelfReport.ps1` (1,396 lines)
- **License:** GPL-3.0
- **Target OS:** Windows 10/11 and Windows Server
- **Repository Size:** Small (single script, no build system)
- **Runtime:** PowerShell 5.1+ on Windows, PowerShell Core 7+ for cross-platform testing

## Project Architecture

### File Structure
```
.
├── AssetSelfReport.ps1    # Main automation script (1,396 lines)
├── README.md              # Setup and configuration documentation
├── LICENSE                # GPL-3.0 license
└── .gitignore            # Excludes config.ps1
```

### Key Components

1. **Configuration** (Lines 1-46): Requires external `config.ps1` file with SnipeIT URL, API key, and local directories
2. **Package/Module Management** (Lines 262-313): Auto-installs NuGet, SnipeitPS, DellBIOSProvider, ActiveDirectory, PSWindowsUpdate
3. **System Information Gathering** (Lines 316-783): Collects hardware, software, BIOS, network, drive, and user data
4. **SnipeIT Integration** (Lines 852-1213): Updates or creates assets in SnipeIT inventory system
5. **Change Detection** (Lines 1241-1322): Compares previous state and sends email alerts on changes
6. **Optional SCCM Integration** (Lines 1338-1381): Triggers SCCM check-ins if enabled

### Dependencies (Auto-installed by script)
- **NuGet** package provider
- **SnipeitPS** module (SnipeIT PowerShell API wrapper)
- **DellBIOSProvider** module (Dell systems only)
- **ActiveDirectory** PowerShell module (RSAT)
- **PSWindowsUpdate** module

### Custom Code Sections
Lines 128-140 contain environment-specific customizations that may cause errors in generic environments. These are marked with `# Begin Custom Code` and `# End Custom Code` comments.

## Execution Requirements

### Prerequisites
1. **PowerShell Version:** 5.1 or later (PowerShell Core 7+ supported but not required)
2. **Administrator Rights:** Script requires elevation for WMI queries, BIOS access, and local group management
3. **Configuration File:** Must create `config.ps1` in the same directory as the script
4. **Network Access:** Requires connectivity to SnipeIT server and Dell warranty API (if used)

### Running the Script

**Command:**
```powershell
pwsh -File ./AssetSelfReport.ps1 -ConfigFile "./config.ps1"
```

**Expected Behavior:**
- Script will output `Write-Error` messages if config file is missing
- Exit code 1 indicates missing configuration
- Script sleeps for 10 seconds before exiting on configuration errors
- Successful execution creates log files in configured directories

### Configuration File Format
The `config.ps1` file should contain JSON data that will be read and parsed by the script using `ConvertFrom-Json`. Create a plain text file with .ps1 extension containing the following JSON structure:
```json
{
  "EmailParams": {
    "From": "sender@domain.com",
    "To": "recipient@domain.com",
    "SMTPServer": "smtp.server.com",
    "Port": 25
  },
  "LocalFileDir": "C:\\Path\\To\\Local\\Files",
  "LogFileDir": "C:\\Path\\To\\Logs",
  "RecordFileDir": "\\\\server\\share\\records",
  "DellApi": {
    "Key": "your-dell-api-key",
    "Secret": "your-dell-api-secret"
  },
  "Snipe": {
    "Url": "https://your.snipeit.url",
    "Token": "your-snipeit-api-token",
    "DefStatusID": 2,
    "FieldSetID": 1,
    "WorkstationCatID": 3,
    "ServerCatID": 4
  },
  "DailyPowerOnList": ["LOC-001", "LOC-002"],
  "DellBios": {
    "KeyFile": "C:\\path\\to\\key.txt",
    "OldPwdFile": "C:\\path\\to\\oldpwd.txt",
    "NewPwdFile": "C:\\path\\to\\newpwd.txt"
  }
}
```

## Testing and Validation

### Syntax Validation
```powershell
# Check PowerShell syntax without execution
pwsh -NoProfile -Command {
    $null = $ExecutionContext.InvokeCommand.NewScriptBlock((Get-Content ./AssetSelfReport.ps1 -Raw))
}
```

### Script Analysis (PowerShell Script Analyzer)
```powershell
# Install PSScriptAnalyzer if not present
Install-Module -Name PSScriptAnalyzer -Force -Scope CurrentUser

# Run analysis
Invoke-ScriptAnalyzer -Path ./AssetSelfReport.ps1 -Severity Error,Warning
```

### Testing Without Windows
- **Linux/macOS:** PowerShell Core is available and can be used for syntax checking only
- **Limitation:** Script requires Windows-specific WMI classes, so functional testing requires Windows
- **Mock Testing:** Not practical due to heavy Windows API dependencies
- **Pester tests:** `tests/` holds Pester 5 tests for pure functions, which are parsed out of `AssetSelfReport.ps1` so the script itself never runs. Run them in a container: `docker run --rm -v "$PWD:/src" -w /src mcr.microsoft.com/powershell:latest pwsh -NoProfile -Command "Install-Module Pester -MinimumVersion 5.0 -Force -Scope CurrentUser; Invoke-Pester ./tests -Output Detailed"`. Keep new logic in small pure functions so it can be tested this way, and add a test for each one.

### Known Validation Issues
1. **Platform Dependency:** Script will fail on non-Windows systems at runtime (WMI classes unavailable)
2. **Configuration Errors:** Missing config file causes immediate exit with error code 1
3. **Custom Code Sections:** Lines 128-140 reference specific hostnames and paths that will fail in other environments

## Common Issues and Workarounds

### Issue 1: Missing Configuration File
**Error:** `Write-Error: You did not provide a config file!`
**Solution:** Always pass `-ConfigFile` parameter with valid JSON config
**Prevention:** Check for config file existence before running script

### Issue 2: Module Installation Failures
**Location:** Lines 262-313 (package and module installation)
**Symptoms:** Errors about missing modules or package providers
**Solution:** 
- Ensure internet connectivity for PowerShell Gallery
- Run with administrator privileges
- Set execution policy: `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass`

### Issue 3: SnipeIT API Errors
**Location:** Lines 852-1213 (SnipeIT integration)
**Symptoms:** `StatusCode='InternalServerError'` messages
**Behavior:** Script retries once automatically (lines 860-865)
**Solution:** Check SnipeIT API token, URL, and network connectivity
**Note:** Script sends email alerts on duplicate assets or search errors

### Issue 4: Dell BIOS Access
**Location:** Lines 345-437 (BIOS configuration)
**Symptoms:** Errors accessing `DellSmbios:\` provider
**Solution:** Script includes try/catch blocks and gracefully handles missing DellBIOSProvider
**Prevention:** Dell-specific code only runs on Dell hardware with `Get-Item -Path "DellSmbios:\"` check

### Issue 5: Custom Code Path Failures
**Location:** Lines 128-140 (custom environment code)
**Symptoms:** Errors about missing paths or specific hostnames
**Solution:** These are environment-specific customizations and can be safely commented out for other environments

## Code Modification Guidelines

### Making Changes
1. **Preserve Custom Fields:** Lines 869-921 define SnipeIT custom field mappings (e.g., `_snipeit_mac_address_1`). These IDs are instance-specific.
2. **Log File Paths:** Changes to directory structures must update Lines 36-46 and 156-162
3. **Software Tracking:** Default software exclusion list (Lines 61-123) should be reviewed when changing software detection
4. **Email Alerts:** All alert functions use `EmailAlert` wrapper (Lines 229-233). Modify SMTP settings in config, not code.

### Testing Changes
1. **Syntax Check:** Use the syntax validation from lines 105-107 of this document
2. **Static Analysis:** Run PSScriptAnalyzer before committing changes
3. **Dry Run:** Test on a non-production machine with test SnipeIT instance
4. **Monitor Logs:** Script creates detailed logs at `$LogFileDir\$DateDir\$DeviceName\`

### Adding New Features
1. **Data Collection:** Add to `$DataHashTable` (hashtable for collected data)
2. **SnipeIT Fields:** Add corresponding entries to `$CustomValues` (lines 869-921)
3. **Change Detection:** Update comparison list in Line 1289 (`$ToCompare` array)
4. **Logging:** Use `WriteLog` function with `-Log` parameter for messages

## CI/CD and Automation

**No CI/CD Pipeline Currently Exists.** This repository has no GitHub Actions workflows, build scripts, or automated tests.

**Recommended Validation Steps:**
1. Create a test `config.ps1` with placeholder values
2. Run syntax validation with PowerShell Core
3. Run PSScriptAnalyzer for best practice violations
4. Test on Windows VM before deploying to production

## Security Considerations

1. **Sensitive Data:** `config.ps1` is excluded via `.gitignore` (contains API keys, SMTP credentials)
2. **Password Handling:** Dell BIOS passwords are stored encrypted and read from files (Lines 352-357)
3. **Local Admin Changes:** Script modifies local Remote Desktop Users group (Lines 971-987)
4. **Email Notifications:** All major changes and errors trigger email alerts
5. **API Tokens:** Never commit SnipeIT API tokens or Dell API credentials

## Quick Reference

### Key Functions
- `WriteLog` (Lines 214-227): Logging with color-coded output
- `EmailAlert` (Lines 229-233): Send email notifications
- `GetHRSize` (Lines 234-244): Human-readable byte sizes
- `CheckFilesAndDirectories` (Lines 250-256): Ensure required paths exist

### Important Variables
- `$DataHashTable`: Collected system information
- `$CustomValues`: SnipeIT custom field mappings
- `$SnipeAsset`: Retrieved SnipeIT asset object
- `$Record`: Previously saved asset state from CSV

### Exit Codes
- `0`: Successful execution
- `1`: Missing configuration file

## Final Notes

**Trust these instructions:** This document has been thoroughly researched by examining the entire codebase. Only perform additional searches if information is incomplete or contradicts what you observe.

**Windows-Only:** This script is fundamentally Windows-dependent. Do not attempt to make it cross-platform without major refactoring.

**Self-Contained:** The script manages its own dependencies through PowerShell module auto-installation. No external build tools required.

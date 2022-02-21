# Asset Self Reporting SnipeIT
A script to compile an assets information and update SnipeIT inventory system.
![Screenshot 2022-02-21 163844](https://user-images.githubusercontent.com/58618324/155029177-dc5c289f-adc7-403b-9619-8528b1267177.png)

## Features:
- Added/Removed Software Alerting. Checks all software, even user profile installed software.
- Significant Configuration Change Alerting. A "significant change" is configured as a change in:
    - DeviceName, IpAddress, MacAddress, NetworkAdapters, CPU, RAM_Installed, Drives, DHCP, OS, Bios, LocalAdmins, RemoteUsers, Graphics, Webcam
- Reports a lot of data that would normally need to be manually entered into SnipeIT.
- Pulls warranty information from Dell API. (Dell API Key required)
- Automated Bios Configurations. (Dell Only)
- Uses device Serial Number to pull existing SnipeIT Asset or create a new one if it does not exist.


## Requirements:
- SnipeIT Inventory System
- Create Custom Fields in SnipeIT for each data point:
    - Mac Address, CPU, RAM, Operating System, IP Address, Bios, Last Reported, Graphics, Boot Drive, Internal Media, External Media, Licensed Software, Remote Desktop Users, Applied Updates, Network Adapters
    - You will need to get the SnipeIT custom field value for each of thsoe custom fields in order to fill out the fields in the config file. 
- In order to run the script, you have to include the parameter `-ConfigFile` with the location of the config file. Example:
        `Powershell.exe -Command "\\path\to\ReportAssetDataToSnipeIT.ps1 -ConfigFile \\path\to\selfReportConfig.json""`
## Recommendations:
- Require all powershell scripts to be signed in your domain environment. Set through Group Policy.
- Sign this script with an organization code signing certificate that is pushed to all domain assets via GPO.
- Run as a scheduled task pushed to all domain assets with a GPO.
- A minimum, set the scheduled task to trigger once per day.

This is a script I use in my environment to automatically update all domain assets daily to my SnipeIT Inventory System and manage certain aspects of the assets based on information the script finds in the inventory system.

This is written to work with Dell machines as I have a Dell exclusive environment. If you're environment includes other manufacturers, you will lose the automated bios configurations and the automated warranty information features. 

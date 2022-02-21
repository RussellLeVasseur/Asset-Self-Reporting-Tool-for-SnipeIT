# Asset Self Reporting SnipeIT
A script to compile an assets information and update SnipeIT inventory system.

Features:
- Added/Removed Software Alerting. Checks all software, even user profile installed software.
- Significant Configuration Change Alerting. A "significant change" is configured as a change in:
    - DeviceName, IpAddress, MacAddress, NetworkAdapters, CPU, RAM_Installed, Drives, DHCP, OS, Bios, LocalAdmins, RemoteUsers, Graphics, Webcam
- Reports all data

Requirements:
- SnipeIT Inventory System
- Create Custom Fields in SnipeIT for each data point:
    - Mac Address, CPU, RAM, Operating System, IP Address, Bios, Last Reported, Graphics, Boot Drive, Internal Media, External Media, Licensed Software, Remote Desktop Users, Applied Updates, Network Adapters

Recommendations:
- Require all powershell scripts to be signed in your domain environment. Set through Group Policy.
- Sign this script with an organization code signing certificate that is pushed to all domain assets via GPO.
- Run as a scheduled task pushed to all domain assets with a GPO.
- A minimum, set the scheduled task to trigger once per day.

This is a script I use in my environment to automatically update all domain assets daily to my SnipeIT Inventory System and manage certain aspects of the assets based on information the script finds in the inventory system.

Sensitive information and some functions have been removed which may cause some errors in a few functions. You will need to modify this for you environment. 

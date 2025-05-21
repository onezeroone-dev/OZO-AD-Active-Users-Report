# OZO AD Active Users Report Installation and Usage

## Description
Produces an Excel report of active domain users including display name, title, office, department, and password status.

## Prerequisites
This script requires the _ActiveDirectory_, _ImportExcel_, _OZO_, and _OZOAD_ PowerShell modules. The _ActiveDirectory_ PowerShell module is included with the [_Remote Server Administration Tools (RSAT) for Windows_](https://learn.microsoft.com/en-us/troubleshoot/windows-server/system-management-components/remote-server-administration-tools) feature installation. The remaining modules are published to [PowerShell Gallery](https://learn.microsoft.com/en-us/powershell/scripting/gallery/overview?view=powershell-5.1). Ensure your system is configured for this repository then execute the following in an _Administrator_ PowerShell:

```powershell
Install-Module ImportExcel,OZO,OZOAD
```

## Installation
This script is published to [PowerShell Gallery](https://learn.microsoft.com/en-us/powershell/scripting/gallery/overview?view=powershell-5.1). Ensure your system is configured for this repository then execute the following in an _Administrator_ PowerShell:

```powershell
Install-Script ozo-ad-active-users-report
```

## Usage
```powershell
ozo-ad-active-users-report
    -Mail   <String>
    -OutDir <String>
```

## Parameters
|Parameter|Description|
|---------|-----------|
|`Mail`|Email address to [attempt to\] send the Excel report to.|
|`OutDir`|Directory for the Excel report. Defaults to the current directory.|

## Outputs
None.

## Notes
Empty cells in the password status field indicates the password has never been set.

## Acknowledgements
Special thanks to my employer, [Sonic Healthcare USA](https://sonichealthcareusa.com), who supports the growth of my PowerShell skillset and enables me to contribute portions of my work product to the PowerShell community.

# Welcome
Microsoft provides an API for programmatic access to security update details using [Common Vulnerability Reporting Format](http://www.icasi.org/cvrf/). View our [blog post](https://msrc-blog.microsoft.com/2016/11/08/furthering-our-commitment-to-security-updates/) for more info. 

The Microsoft [Security Update Guide](https://msrc.microsoft.com/update-guide) is the web experience to find security update detail.

This repository contains sample code and documentation for the Microsoft Security Updates API (https://portal.msrc.microsoft.com/en-us/developer), including:
* source code for the [MsrcSecurityUpdates PowerShell module](https://www.powershellgallery.com/packages/MsrcSecurityUpdates)
* sample code for using the [MsrcSecurityUpdates PowerShell module](https://www.powershellgallery.com/packages/MsrcSecurityUpdates)
* OpenAPI/Swagger definition for the Microsoft Security Updates API

# Getting the MsrcSecurityUpdates PowerShell Module
Getting started with the MsrcSecurityUpdates module can be done like this:
```PowerShell
### Install the module from the PowerShell Gallery
###  !! Requires minimum PowerShell version 5.1
Install-Module -Name MsrcSecurityUpdates -Scope CurrentUser

### Load the module
Import-Module -Name MsrcSecurityUpdates
```
Once the module is loaded, check out our [PowerShell samples](https://github.com/Microsoft/MSRC-Microsoft-Security-Updates-API/blob/master/src/README.md)

# API Keys
The Security Updates API requires an API key.  To obtain an API key please visit the [Security Update Guide Developer page](https://portal.msrc.microsoft.com/en-us/developer).  For help using the Security Updates Guide please visit the [Security Updates Guide Community Forum](https://social.technet.microsoft.com/Forums/security/en-us/home?forum=securityupdateguide).

__NOTE: Generating an API key requires signing in with an @outlook.com, @live.com, or @microsoft.com email address. If you do not have one of these email addresses, you can create a personal outlook account to access this service. In the future, we will be removing this authentication requirement entirely.__

# Change Log
**_For up to date major changes, please read the psd1 included in the src folder. This can also be seen on [the Microsoft Powershell Gallery](https://www.powershellgallery.com/packages/MsrcSecurityUpdates)._**

# Support
## Developer Support
Customers should treat this repository as custom code.  Bug fixes or enhancements can be requested by opening a new issue from the Issues tab.
## Security Update Support
For questions about CVEs, security updates and patches, please visit [Microsoft Support](https://support.microsoft.com)
## Security Update Guide Support
For questions about the [Microsoft Security Update Guide](https://msrc.microsoft.com/update-guide) please visit the [Security Update Guide support forum](https://social.technet.microsoft.com/Forums/security/en-us/home?forum=securityupdateguide).

# Contributing

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

It is Microsoft’s mission to empower every person and every organization on the planet to achieve more. We thank you for helping shape that future by keeping the world a more secure place by tooling security into your organization’s practices. We would love to hear your feedback on features to add or bugs to fix.

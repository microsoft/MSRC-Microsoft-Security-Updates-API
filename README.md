# Welcome
Microsoft provides an API for getting security update details using [Common Vulnerability Reporting Format](http://www.icasi.org/cvrf/). View our [blog post](https://blogs.technet.microsoft.com/msrc/2016/11/08/furthering-our-commitment-to-security-updates/) for more info. 

The [Security Updates Guide](https://portal.msrc.microsoft.com/en-us/security-guidance) is a great place to find security updates in a browser, and the Security Updates API is intended for doing automation around Microsoft security updates.

This project contains sample code and documentation for the Microsoft Security Updates API (https://portal.msrc.microsoft.com/en-us/developer), including:
* source code for the [MsrcSecurityUpdates PowerShell module](https://www.powershellgallery.com/packages/MsrcSecurityUpdates)
* sample code for using the [MsrcSecurityUpdates PowerShell module](https://www.powershellgallery.com/packages/MsrcSecurityUpdates)
* OpenAPI/Swagger definition for the Microsoft Security Updates API

# Getting the MsrcSecurityUpdates PowerShell Module
Getting started with the MsrcSecurityUpdates module can be done like this:
```PowerShell
### Install the module from the PowerShell Gallery
###  !! Requires PowerShell V5
###  !! Install-Module requires admin permission
Install-Module -Name MsrcSecurityUpdates

### Load the module
Import-Module -Name MsrcSecurityUpdates
```
Once the module is loaded, check out our [PowerShell samples](https://github.com/Microsoft/MSRC-Microsoft-Security-Updates-API/blob/master/src/README.md)

# API Keys
The Security Updates API requires an API key.  To obtain an API key please visit the [Security Updates Guide](https://portal.msrc.microsoft.com/en-us/security-guidance).  For help using the Security Updates Guide please visit the [Security Updates Guide Community Forum](https://social.technet.microsoft.com/Forums/security/en-us/home?forum=securityupdateguide).

__NOTE: Currently generating api keys requires an @outlook.com, @live.com, or @microsoft.com email address. If you do not have one of these email addresses, you can create a personal outlook account to access this service while we investigate this issue.__

# Change Log
**_For up to date major changes, please read the psd1 included in the src folder. This can also be seen on [the Microsoft Powershell Gallery](https://www.powershellgallery.com/packages/MsrcSecurityUpdates)._**

**May 2017** - Added RestartRequired and SubType to the remediations object in the API response. 


# Contributing

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

It is Microsoft’s mission to empower every person and every organization on the planet to achieve more. We thank you for helping shape that future by keeping the world a more secure place by tooling security into your organization’s practices. We would love to hear your feedback on features to add or bugs to fix.

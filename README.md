# Welcome
This project serves to help developers quickly start processing data from the Microsoft Security Updates API (https://portal.msrc.microsoft.com/en-us/developer)

It is comprised of two main components, documentation and sample code. You'll find OpenAPI/Swagger definitions in the docs folder, which you can use to generate your own client via swagger.io.  

If you are a PowerShell developer, you can install PowerShell cmdlets directly by adding ``import module MSRCsecurityupdates`` to your PowerShell script, or download the source code from this repo and tailor it according to your requirements. These PowerShell cmdlets abstract CVRF data into data structure tailored for common reporting & automation scenarios. 

To see information about these, please view the README.md in the relevant folders.

# Background
Security bulletins are being replaced with the industry standard CVRF report type. View our [blog post](https://blogs.technet.microsoft.com/msrc/2016/11/08/furthering-our-commitment-to-security-updates/) for more info. Please use the [Security Updates Guide](https://portal.msrc.microsoft.com/en-us/security-guidance) to view vulnerabilities, or use this project to help automate Microsoft's vulnerability reporting & automation within your organization.

# Change Log

**March 14, 2017** – Minor changes to Powershell module to fix a CVRF->Powershell object conversion issue. Republished new Powershell module to [the Microsoft Powershell Gallery](https://www.powershellgallery.com/packages/MsrcSecurityUpdates/1.2). 

**March 09, 2017** – Added revised PowerShell cmdlets. 

**February 09, 2017** - Project is going live.



# Contributing

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

It is Microsoft’s mission to empower every person and every organization on the planet to achieve more. We thank you for helping shape that future by keeping the world a more secure place by tooling security into your organization’s practices. We would love to hear your feedback on features to add or bugs to fix.

# Sample Code

The sample code serves as an example on how to interact with the MSRC Security Updates API through Powershell.  

In the PowerShell module, you will see script functions that show how to interact with the API, as well 
as functions that organize data from the industry-standard [CVRF document format](http://www.icasi.org/cvrf-v1-1-dictionary-of-elements/#40rem)  to show some reporting scenarios.

You can clone this repo, or download the latest modules directly from the Microsoft Powershell Gallery (Powershell v3 and higher). 

## Installing directly from Microsoft Powershell Gallery

**Must be running Powershell v5.1 or higher** 
With this solution, you do not need to download code from this Github repository. 
Instead, you can download the directly from the Microsoft Powershell Gallery.
The Powershell version can be verified by checking the Major variable in the ``$PSVersionTable.PSVersion`` variable:

```
$PSVersionTable.PSVersion

Major  Minor  Build  Revision
-----  -----  -----  --------
5      1      14393  693     
```

If you're running **as an administrator**  you can install the module in your program files:
```Powershell
Install-Module MSRCSecurityUpdates -Force 
Import-Module MSRCSecurityUpdates
````

If you're **not running as an administrator**, you can install the module in your profile by providing an additional parameter about the Scope:
```Powershell
Install-Module MSRCSecurityUpdates -Force -Scope CurrentUser
Import-Module MSRCSecurityUpdates
````

## Hello, world!
See *MsrcSecurityUpdates.tests.ps1* which exercises all of the functions.

## Cmdlets

```Powershell
Get-Command -Module MsrcSecurityUpdates

CommandType     Name                                               Version    Source                                                                                                 
-----------     ----                                               -------    ------                                                                                                 
Function        Get-KBDownloadUrl                                  1.8.7      MsrcSecurityUpdates
Function        Get-MsrcCvrfAffectedSoftware                       1.8.7      MsrcSecurityUpdates
Function        Get-MsrcCvrfCVESummary                             1.8.7      MsrcSecurityUpdates
Function        Get-MsrcCvrfDocument                               1.8.7      MsrcSecurityUpdates
Function        Get-MsrcCvrfExploitabilityIndex                    1.8.7      MsrcSecurityUpdates
Function        Get-MsrcSecurityBulletinHtml                       1.8.7      MsrcSecurityUpdates
Function        Get-MsrcSecurityUpdate                             1.8.7      MsrcSecurityUpdates
Function        Get-MsrcVulnerabilityReportHtml                    1.8.7      MsrcSecurityUpdates
Function        Set-MSRCAdalAccessToken                            1.8.7      MsrcSecurityUpdates
Function        Set-MSRCApiKey                                     1.8.7      MsrcSecurityUpdates
```

## Generating a HTML document of Monthly Updates

In this common scenario, the *Get-MsrcCvrfDocument* and *Get-MsrcVulnerabilityReportHtml* can be pipelined together to generate a HTML document with *out-file*:

```Powershell
### Install the module from the PowerShell Gallery (must be run as Admin)
Install-Module -Name msrcsecurityupdates -force
Import-module MsrcSecurityUpdates
$monthOfInterest = '2017-Apr'

Get-MsrcCvrfDocument -ID $monthOfInterest -Verbose | 
Get-MsrcSecurityBulletinHtml -Verbose | 
Out-File c:\temp\MSRCAprilSecurityUpdates.html
```
You can also build a modified object to pass into *Get-MsrcVulnerabilityReportHtml*. This allows more customization of the report being generated. In this example, the generated report will only have the wanted CVE's included:

```Powershell
Install-Module -Name MsrcSecurityUpdates -Force
Import-Module -Name MsrcSecurityUpdates -Force

$monthOfInterest = "2017-Mar"

$CVEsWanted = @(
        "CVE-2017-0001", 
        "CVE-2017-0005"
        )
$Output_Location = "C:\your\path\here"

$CVRFDoc = Get-MsrcCvrfDocument -ID $monthOfInterest -Verbose
$CVRFHtmlProperties = @{
    Vulnerability = $CVRFDoc.Vulnerability | Where-Object {$_.CVE -in $CVEsWanted}
    ProductTree = $CVRFDoc.ProductTree
}

Get-MsrcVulnerabilityReportHtml @CVRFHtmlProperties -Verbose | Out-File $Output_Location
```

An alternate HTML template is also avalible which generates reports which are grouped into catagories rather than by each CVE. This report may be more helpful if you are interested in all the vulnerabilities that affect a certain product and platform. This report can be ran exactly like above, however if you are filtering based on CVE's (or not piping the output from *Get-MsrcCvrfDocument*) two aditional fields are required. the above examples are replicated here for the different report type.

Building a report that contains all CVE's:

```Powershell
### Install the module from the PowerShell Gallery (must be run as Admin)
Install-Module -Name MsrcSecurityUpdates -Force
Import-module MsrcSecurityUpdates
$monthOfInterest = '2017-Apr'

Get-MsrcCvrfDocument -ID $monthOfInterest -Verbose | Get-MsrcSecurityBulletinHtml -Verbose | Out-File c:\temp\MSRCAprilSecurityUpdates.html
```

Using powershell to filter the report to your liking:

```Powershell
Install-Module -Name MsrcSecurityUpdates -Force
Import-Module -Name MsrcSecurityUpdates -Force

$monthOfInterest = "2017-Mar"

$CVEsWanted = @(
        "CVE-2017-0001", 
        "CVE-2017-0005"
        )
$Output_Location = "C:\your\path\here"

$CVRFDoc = Get-MsrcCvrfDocument -ID $monthOfInterest -Verbose
$CVRFHtmlProperties = @{
    Vulnerability = $CVRFDoc.Vulnerability | Where-Object {$_.CVE -in $CVEsWanted}
    ProductTree = $CVRFDoc.ProductTree
    DocumentTracking = $CVRFDoc.DocumentTracking
    DocumentTitle = $CVRFDoc.DocumentTitle
}

Get-MsrcSecurityBulletinHtml @CVRFHtmlProperties -Verbose | Out-File $Output_Location
```

## Finding Mitigations and Workarounds

In this scenario, you can use the *Get-MsrcCvrfDocument* and extract the migrations & remediations from each vulnerability.

```Powershell
### Install the module from the PowerShell Gallery (must be run as Admin)
Install-Module -Name MsrcSecurityUpdates

### Download the March CVRF as an object
$cvrfDoc = Get-MsrcCvrfDocument -ID 2017-Mar

### Get the Remediations of Type 'Workaround' (0)
$cvrfDoc.Vulnerability.Remediations | Where Type -EQ 0

### Get the Remediations of Type 'Mitigation' (1)
$cvrfDoc.Vulnerability.Remediations | Where Type -EQ 1

### Get the Remediations of Type 'VendorFix' (2)
$cvrfDoc.Vulnerability.Remediations | Where Type -EQ 2 
```

## Help!
You'll find help via ``Get-Help`` for each cmdlet:

```Powershell
Get-Help Get-MsrcSecurityUpdate

NAME
    Get-MsrcSecurityUpdate

SYNOPSIS
    Get MSRC security updates


SYNTAX
    Get-MsrcSecurityUpdate [<CommonParameters>]

    Get-MsrcSecurityUpdate [-After <DateTime>] [-Before <DateTime>] [<CommonParameters>]

    Get-MsrcSecurityUpdate -Year <Int32> [<CommonParameters>]

    Get-MsrcSecurityUpdate -Vulnerability <String> [<CommonParameters>]


DESCRIPTION
    Calls the CVRF Update API to get a list of security updates


RELATED LINKS

REMARKS
    To see the examples, type: "get-help Get-MsrcSecurityUpdate -examples".
    For more information, type: "get-help Get-MsrcSecurityUpdate -detailed".
    For technical information, type: "get-help Get-MsrcSecurityUpdate -full".
```

... as well as sample code with ``--examples``:

```Powershell
Get-Help Get-MsrcSecurityUpdate -examples

NAME
    Get-MsrcSecurityUpdate

SYNOPSIS
    Get MSRC security updates

    -------------------------- EXAMPLE 1 --------------------------

    PS C:\>Get-MsrcSecurityUpdate


    Get all the updates




    -------------------------- EXAMPLE 2 --------------------------

    PS C:\>Get-MsrcSecurityUpdate -Vulnerability CVE-2017-0003


    Get all the updates containing Vulnerability CVE-2017-003




    -------------------------- EXAMPLE 3 --------------------------

    PS C:\>Get-MsrcSecurityUpdate -Year 2017


    Get all the updates for the year 2017




    -------------------------- EXAMPLE 4 --------------------------

    PS C:\>Get-MsrcSecurityUpdate -Cvrf 2017-Jan


    Get all the updates for the CVRF document with ID of 2017-Jan




    -------------------------- EXAMPLE 5 --------------------------

    PS C:\>Get-MsrcSecurityUpdate -Before 2017-01-01


    Get all the updates before January 1st, 2017




    -------------------------- EXAMPLE 6 --------------------------

    PS C:\>Get-MsrcSecurityUpdate -After 2017-01-01


    Get all the updates after January 1st, 2017




    -------------------------- EXAMPLE 7 --------------------------

    PS C:\>Get-MsrcSecurityUpdate -Before 2017-01-01 -After 2016-10-01


    Get all the updates before January 1st, 2017 and after October 1st, 2016




    -------------------------- EXAMPLE 8 --------------------------

    PS C:\>Get-MsrcSecurityUpdate -After (Get-Date).AddDays(-60) -Before (Get-Date)


    Get all updates between now and the last 60 days

```

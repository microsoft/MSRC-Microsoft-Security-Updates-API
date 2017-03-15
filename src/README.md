# Sample Code

The sample code serves as an example on how to interact with the MSRC Security Updates API through Powershell.  You will need to log into the MSRC Portal and obtain an API key.   

In the PowerShell module, you will see script functions that show how to interact with the API, as well 
as functions that organize data from the industry-standard CVRF document format to show some reporting scenarios.

You can clone this repo, or download the latest modules directly from the Microsoft Powershell Gallery (Powershell v3 and higher). 

## Installing directly from Microsoft Powershell Gallery

**Must be run as Administrator and running Powershell v3 or higher**  The Powershell version can be verified by checking the Major variable in the ``$psversiontable.PSVersion``  variable:

```
$psversiontable.PSVersion

Major  Minor  Build  Revision
-----  -----  -----  --------
5      1      14393  693     
```

Once you meet the above conditions, add the following to your Powershell script:

```Powershell
Install-Module MSRCSecurityUpdates -force
````


## Hello, world!
See *MsrcSecurityUpdates.tests.ps1* which exercises all of the functions.


## Cmdlets

```Powershell
get-command -Module MsrcSecurityUpdates

CommandType     Name                                               Version    Source
-----------     ----                                               -------    ------
Function        Get-MsrcCvrfDocument                               1.1        MsrcSecurityUpdates
Function        Get-MsrcCvrfProductVulnerability                   1.1        MsrcSecurityUpdates
Function        Get-MsrcSecurityBulletinHtml                       1.1        MsrcSecurityUpdates
Function        Get-MsrcSecurityUpdate                             1.1        MsrcSecurityUpdates
```

## Generating a HTML document of Monthly Updates

In this common scenario, the *Get-MsrcCvrfDocument* and *Get-MsrcSecurityBulletinHtml* can be pipelined together to generate a HTML document with *out-file*:

```Powershell
$msrcAPIKey = "<your API key>"
$monthOfInterest = "2016-Nov"

Get-MsrcCvrfDocument -ID $monthOfInterest -ApiKey $msrcApiKey -Verbose | Get-MsrcSecurityBulletinHtml -Verbose | Out-File c:\temp\MSRCNovSecurityUpdates.html
```


## Help!
You'll find help via ``get-help`` for each cmdlet:

```Powershell
get-help Get-MsrcSecurityUpdate

NAME
    Get-MsrcSecurityUpdate
    
SYNOPSIS
    Get MSRC security updates
    
    
SYNTAX
    Get-MsrcSecurityUpdate -ApiKey <String> [<CommonParameters>]
    
    Get-MsrcSecurityUpdate -ApiKey <String> -Vulnerability <String> [<CommonParameters>]
    
    Get-MsrcSecurityUpdate -ApiKey <String> -Cvrf <String> [<CommonParameters>]
    
    Get-MsrcSecurityUpdate -ApiKey <String> -Year <Int32> [<CommonParameters>]
    
    Get-MsrcSecurityUpdate -ApiKey <String> [-After <DateTime>] [-Before <DateTime>] [<CommonParameters>]
    
    
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
get-help Get-MsrcSecurityUpdate -examples

NAME
    Get-MsrcSecurityUpdate
    
SYNOPSIS
    Get MSRC security updates
    
    -------------------------- EXAMPLE 1 --------------------------
    
    PS C:\>#Get all the updates
    Get-MsrcSecurityUpdate -ApiKey 'YOUR API KEY'
    
    -------------------------- EXAMPLE 2 --------------------------
    
    PS C:\>#Get all the updates containing Vulnerability CVE-2017-003
    Get-MsrcSecurityUpdate -ApiKey 'YOUR API KEY' -Vulnerability CVE-2017-0003
    
    -------------------------- EXAMPLE 3 --------------------------
    
    PS C:\>#Get all the updates for the year 2017
    Get-MsrcSecurityUpdate -ApiKey 'YOUR API KEY' -Year 2017
    
    -------------------------- EXAMPLE 4 --------------------------
    
    PS C:\>#Get all the updates for the CVRF document with ID of 2017-Jan
    Get-MsrcSecurityUpdate -ApiKey 'YOUR API KEY' -Cvrf 2017-Jan
    
    -------------------------- EXAMPLE 5 --------------------------
    
    PS C:\>#Get all the updates before January 1st, 2017
    Get-MsrcSecurityUpdate -ApiKey 'YOUR API KEY' -Before 2017-01-01
    
    -------------------------- EXAMPLE 6 --------------------------
    
    PS C:\>#Get all the updates after January 1st, 2017
    Get-MsrcSecurityUpdate -ApiKey 'YOUR API KEY' -After 2017-01-01
    
    -------------------------- EXAMPLE 7 --------------------------
    
    PS C:\>#Get all the updates before January 1st, 2017 and after October 1st, 2016
    Get-MsrcSecurityUpdate -ApiKey 'YOUR API KEY' -Before 2017-01-01 -After 2016-10-01
```
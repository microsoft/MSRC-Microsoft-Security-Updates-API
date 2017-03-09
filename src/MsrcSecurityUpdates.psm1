$msrcApiUrl     = 'https://api.msrc.microsoft.com'
$msrcApiVersion = 'api-version=2016-08-01'

function Get-MsrcSecurityUpdate
{
<#
.Synopsis
   Get MSRC security updates
.DESCRIPTION
   Calls the CVRF Update API to get a list of security updates
.EXAMPLE
   #Get all the updates
   Get-MsrcSecurityUpdate -ApiKey 'YOUR API KEY'
.EXAMPLE
   #Get all the updates containing Vulnerability CVE-2017-003
   Get-MsrcSecurityUpdate -ApiKey 'YOUR API KEY' -Vulnerability CVE-2017-0003
.EXAMPLE
   #Get all the updates for the year 2017
   Get-MsrcSecurityUpdate -ApiKey 'YOUR API KEY' -Year 2017
.EXAMPLE
   #Get all the updates for the CVRF document with ID of 2017-Jan
   Get-MsrcSecurityUpdate -ApiKey 'YOUR API KEY' -Cvrf 2017-Jan
.EXAMPLE
   #Get all the updates before January 1st, 2017
   Get-MsrcSecurityUpdate -ApiKey 'YOUR API KEY' -Before 2017-01-01
.EXAMPLE
   #Get all the updates after January 1st, 2017
   Get-MsrcSecurityUpdate -ApiKey 'YOUR API KEY' -After 2017-01-01
.EXAMPLE
   #Get all the updates before January 1st, 2017 and after October 1st, 2016
   Get-MsrcSecurityUpdate -ApiKey 'YOUR API KEY' -Before 2017-01-01 -After 2016-10-01
#>
    [CmdletBinding(DefaultParametersetName="AllUpdates")]      
    Param
    (
        <#
        API Key for the MSRC CVRF API
        To get an API key, visit https://portal.msrc.microsoft.com
        #> 
        [Parameter(Mandatory=$true,ParameterSetName='AllUpdates')]
        [Parameter(Mandatory=$true,ParameterSetName='ByDate')]
        [Parameter(Mandatory=$true,ParameterSetName='ByYear')]
        [Parameter(Mandatory=$true,ParameterSetName='ByCVRF')]
        [Parameter(Mandatory=$true,ParameterSetName='ByVulnerability')]
        [String]
        $ApiKey,

        <#
        Get security updates released after this date
        #>
        [Parameter(ParameterSetName='ByDate')]
        [DateTime]
        $After,

        <#
        Get security updates released before this date
        #>
        [Parameter(ParameterSetName='ByDate')]
        [DateTime]
        $Before,

        <#
        Get security updates for the specified year (ie. 2016)
        #>
        [Parameter(Mandatory=$true,ParameterSetName='ByYear')]        
        [ValidateScript({
            if ($_ -lt 2016 -or $_ -gt [DateTime]::Now.Year) 
            {
                throw 'Year must be between 2016 and this year'
            }
            else
            {
                $true
            }
        })] 
        [Int]
        $Year,

        <#
        Get security updates for the specified Vulnerability CVE (ie. CVE-2016-0128)
        #>
        [Parameter(Mandatory=$true,ParameterSetName='ByVulnerability')]
        [String]
        $Vulnerability,

        <#
        Get security update for the specified CVRF ID (ie. 2016-Aug)
        #>
        [Parameter(Mandatory=$true,ParameterSetName='ByCVRF')]
        [String]
        $Cvrf
    )
    ### Construct the URL based on the parameters provided
    switch ($PSCmdlet.ParameterSetName)
    {
        ByDate {
                if ($PSBoundParameters.ContainsKey('Before') -and $PSBoundParameters.ContainsKey('After'))
                {                                       
                    $url = "$msrcApiUrl/Updates?`$filter=CurrentReleaseDate gt {0} and CurrentReleaseDate lt {1}&$msrcApiVersion" -f $After.ToString('yyyy-MM-dd'), $Before.ToString('yyyy-MM-dd')
                }
                elseif ($PSBoundParameters.ContainsKey('Before'))
                {
                    $url = "$msrcApiUrl/Updates?`$filter=CurrentReleaseDate lt {0}&$msrcApiVersion" -f $Before.ToString('yyyy-MM-dd')
                }
                elseif ($PSBoundParameters.ContainsKey('AFter'))
                {
                    $url = "$msrcApiUrl/Updates?`$filter=CurrentReleaseDate gt {0}&$msrcApiVersion" -f $After.ToString('yyyy-MM-dd')
                }
                else
                {
                    throw 'Unexpected parameter set'
                }
            }
        ByYear {
                $url = "$msrcApiUrl/Updates('$Year')?$msrcApiVersion"
            }
        ByVulnerability {
                $url = "$msrcApiUrl/Updates('$Vulnerability')?$msrcApiVersion"
            }
        ByCVRF {
                $url = "$msrcApiUrl/Updates('$Cvrf')?$msrcApiVersion"
            }
        Default {
                $url = "$msrcApiUrl/Updates?$msrcApiVersion"
            }
    }

    try
    {        
        Write-Verbose "Calling $url"
        $webResponse = Invoke-RestMethod -Uri $url -Headers @{
            'Accept'  = 'application/json'
            'Api-Key' = $ApiKey
        }

        if (-not $webResponse)
        {
            Write-Warning "No results returned from the /Update API"
            return
        }

        Write-Output $webResponse.Value
    } 
    catch 
    {
        Write-Error "HTTP Get failed with status code $($_.Exception.Response.StatusCode): $($_.Exception.Response.StatusDescription)"       
    }
}

function Get-MsrcCvrfDocument
{
<#
.Synopsis
   Get a MSRC CVRF document
.DESCRIPTION
   Calls the MSRC CVRF API to get a CVRF document by ID
.EXAMPLE
   #Get the Cvrf document '2016-Aug' (returns an object)
   Get-MsrcCvrfDocument -ID 2016-Aug -ApiKey 'YOUR API KEY'
.EXAMPLE
   #Get the Cvrf document '2016-Aug' (returns an XML string)
   Get-MsrcCvrfDocument -ID 2016-Aug -ApiKey 'YOUR API KEY' -AsXml
.EXAMPLE
   #Get the Cvrf document '2016-Aug' (returns a JSON string)
   Get-MsrcCvrfDocument -ID 2016-Aug -ApiKey 'YOUR API KEY' -AsJson

#>   
    [CmdletBinding(DefaultParametersetName="DefaultCvrfParameterSet")]     
    Param
    (
        <#
        API Key for the MSRC CVRF API
        To get an API key, visit https://portal.msrc.microsoft.com
        #> 
        [Parameter(Mandatory=$true,ParameterSetName='DefaultCvrfParameterSet')]
        [Parameter(Mandatory=$true,ParameterSetName='XmlOutput')]
        [Parameter(Mandatory=$true,ParameterSetName='JsonOutput')]
        [String]
        $ApiKey,

        <#
        Get the CVRF document for the specified CVRF ID (ie. 2016-Aug)
        #>
        [Parameter(Mandatory=$true,ParameterSetName='DefaultCvrfParameterSet')]
        [Parameter(Mandatory=$true,ParameterSetName='XmlOutput')]
        [Parameter(Mandatory=$true,ParameterSetName='JsonOutput')]
        [String]
        $ID,

        <#
        Output as an XML string
        #>
        [Parameter(Mandatory=$true,ParameterSetName='XmlOutput')]
        [Switch]
        $AsXml,

        <#
        Output as an XML string
        #>
        [Parameter(Mandatory=$true,ParameterSetName='JsonOutput')]
        [Switch]
        $AsJson
    )

    $url = "$msrcApiUrl/cvrf/{0}?$msrcApiVersion" -f $ID

    try
    {        
        Write-Verbose "Calling $url"

        switch ($PSCmdlet.ParameterSetName)
        {
            DefaultCvrfParameterSet {
                Invoke-RestMethod -Uri $url -Headers @{
                    'Api-Key' = $ApiKey
                    'Accept'  = 'application/json'
                }
            }
            XmlOutput {
                $webResponse = Invoke-WebRequest -Uri $url -Headers @{
                    'Api-Key' = $ApiKey
                    'Accept'  = 'application/xml'
                }
                Write-Output $webResponse.Content
            }
            JsonOutput {
                $webResponse = Invoke-WebRequest -Uri $url -Headers @{
                    'Api-Key' = $ApiKey
                    'Accept'  = 'application/json'
                }
                Write-Output $webResponse.Content
            }
        }        
    } 
    catch 
    {
        Write-Error "HTTP Get failed with status code $($_.Exception.Response.StatusCode): $($_.Exception.Response.StatusDescription)"       
    }
}

function Get-MsrcCvrfProductVulnerability
{
<#
.Synopsis
   Get product vulnerability details from a CVRF document
.DESCRIPTION
   CVRF documents next products into several places, including:
   -Vulnerabilities
   -Threats
   -Remediations
   -Product Tree
   This function gathers the details for each product identified in a CVRF document.
   It provides a list of Threats, Remediations and CVSS Score Sets for each product.

.EXAMPLE
   #Get product vulnerability details from a CVRF document using the pipeline   
   Get-MsrcCvrfDocument -ID 2016-Nov -ApiKey 'YOUR API KEY' | Get-MsrcCvrfProductVulnerability
.EXAMPLE
   #Get product vulnerability details from a CVRF document using a variable and parameters
   $cvrfDocument = Get-MsrcCvrfDocument -ID 2016-Nov -ApiKey 'YOUR API KEY'
   Get-MsrcCvrfProductVulnerability -Vulnerability $cvrfDocument.Vulnerability -ProductTree $cvrfDocument.ProductTree -DocumentTracking $cvrfDocument.DocumentTracking -DocumentTitle $cvrfDocument.DocumentTitle
#>        
    Param
    (
        <#
        API Key for the MSRC CVRF API
        To get an API key, visit https://portal.msrc.microsoft.com
        #> 
        [Parameter(Mandatory=$true, 
                   ValueFromPipelineByPropertyName=$true)]
        $Vulnerability,

        <#
        Get the CVRF document for the specified CVRF ID (ie. 2016-Aug)
        #>
        [Parameter(Mandatory=$true, 
                   ValueFromPipelineByPropertyName=$true)]
        $ProductTree,

        <#
        Get the CVRF document for the specified CVRF ID (ie. 2016-Aug)
        #>
        [Parameter(Mandatory=$true, 
                   ValueFromPipelineByPropertyName=$true)]
        $DocumentTracking,

        <#
        Get the CVRF document for the specified CVRF ID (ie. 2016-Aug)
        #>
        [Parameter(Mandatory=$true, 
                   ValueFromPipelineByPropertyName=$true)]
        $DocumentTitle
    )

    ### Create a list of all the products in the CVRF document
    $CvrfRelatedProducts = @()
    foreach ($branch in $ProductTree.Branch.Items)
    {
        foreach ($branchItem in $branch.Items)
        {
            $CvrfRelatedProducts += [psCustomObject]@{        
                CvrfAlias   = $DocumentTracking.Identification.Alias.Value
                CvrfTitle   = $DocumentTitle.Value
                BranchName  = $branch.Name
                ProductName = $branchItem.Value
                ProductID   = $branchItem.ProductID
            }
        }
    }

    ### For each product, get the:
    ###  Threats
    ###  Remediations
    ###  CVSS Score Sets
    foreach ($CvrfRelatedProduct in $CvrfRelatedProducts)
    {
        $Remediations  = @()
        $Threats       = @()
        $CVSSScoreSets = @()
        foreach ($vuln in $Vulnerability)
        {    
            foreach ($Remediation in $vuln.Remediations | Where-Object ProductID -Contains $CvrfRelatedProduct.ProductID)
            {
                $Remediation | Add-Member -NotePropertyName VulnerabilityCVE   -NotePropertyValue $vuln.CVE -Force
                $Remediation | Add-Member -NotePropertyName VulnerabilityTitle -NotePropertyValue $vuln.Title.Value  -Force  
                $Remediations +=  $Remediation    
            }
            $CvrfRelatedProduct | Add-Member -NotePropertyName Remediations -NotePropertyValue $Remediations -Force

            foreach ($Threat in $vuln.Threats | Where-Object ProductID -Contains $CvrfRelatedProduct.ProductID)
            {
                $Threat | Add-Member -NotePropertyName VulnerabilityCVE   -NotePropertyValue $vuln.CVE -Force
                $Threat | Add-Member -NotePropertyName VulnerabilityTitle -NotePropertyValue $vuln.Title.Value -Force
                $Threats += $Threat 
            }
            $CvrfRelatedProduct | Add-Member -NotePropertyName Threats -NotePropertyValue $Threats -Force

            foreach ($CVSSScoreSet in $vuln.CVSSScoreSets | Where-Object ProductID -Contains $CvrfRelatedProduct.ProductID)
            {
                $CVSSScoreSet | Add-Member -NotePropertyName VulnerabilityCVE   -NotePropertyValue $vuln.CVE -Force
                $CVSSScoreSet | Add-Member -NotePropertyName VulnerabilityTitle -NotePropertyValue $vuln.Title.Value -Force
                $CVSSScoreSets += $CVSSScoreSet 
            }            

            $CvrfRelatedProduct | Add-Member -NotePropertyName CVSSScoreSets -NotePropertyValue $CVSSScoreSets -Force
        }
        #region MaximumSeverity
        $MaximumSeverity = 'Unknown'
        $SeverityValues = $Vulnerability.Threats | 
          Where-Object ProductID -Contains $CvrfRelatedProduct.ProductID |
          Where-Object Type -EQ 3 | 
          Select @{Name='Severity' ;Expression={$_.Description.Value}} -Unique |
          Select -ExpandProperty Severity
        
        if ($SeverityValues -contains 'Critical')
        {
            $MaximumSeverity = 'Critical'
        }
        elseif ($SeverityValues -contains 'Important')
        {
            $MaximumSeverity = 'Important'
        }
        elseif ($SeverityValues -contains 'Moderate')
        {
            $MaximumSeverity = 'Moderate'
        }
        elseif ($SeverityValues -contains 'Low')
        {
            $MaximumSeverity = 'Low'
        }
        else
        {
            Write-Verbose "Could not determine the Maximum Severity from the Threats"
        }
        $CvrfRelatedProduct | Add-Member -NotePropertyName MaximumSeverity -NotePropertyValue $MaximumSeverity -Force
        #endregion 

        #region RestartRequired
        $RestartRequired = 'Unknown'
        if ($Vulnerability.Remediations.RestartRequired.Value -contains 'Yes')
        {
            $RestartRequired = 'Yes'
        }
        elseif ($Vulnerability.Remediations.RestartRequired.Value -contains 'Maybe')
        {
            $RestartRequired = 'Maybe'
        }
        $CvrfRelatedProduct | Add-Member -NotePropertyName RestartRequired -NotePropertyValue $RestartRequired -Force
        #endregion

        $CvrfRelatedProduct
    }
}

function Get-MsrcSecurityBulletinHtml
{
#TODO - refactor the code used for populating the tables into functions
    Param
    (
        <#
        API Key for the MSRC CVRF API
        To get an API key, visit https://portal.msrc.microsoft.com
        #> 
        [Parameter(Mandatory=$true, 
                   ValueFromPipelineByPropertyName=$true)]
        $Vulnerability,

        <#
        Get the CVRF document for the specified CVRF ID (ie. 2016-Aug)
        #>
        [Parameter(Mandatory=$true, 
                   ValueFromPipelineByPropertyName=$true)]
        $ProductTree,

        <#
        Get the CVRF document for the specified CVRF ID (ie. 2016-Aug)
        #>
        [Parameter(Mandatory=$true, 
                   ValueFromPipelineByPropertyName=$true)]
        $DocumentTracking,

        <#
        Get the CVRF document for the specified CVRF ID (ie. 2016-Aug)
        #>
        [Parameter(Mandatory=$true, 
                   ValueFromPipelineByPropertyName=$true)]
        $DocumentTitle
    )    

    $htmlDocumentTemplate = @'
<html>
<head>
    <link rel="stylesheet" href="https://i-technet.sec.s-msft.com/Combined.css?resources=0:ImageSprite,0:TopicResponsive,0:TopicResponsive.MediaQueries,1:CodeSnippet,1:ProgrammingSelector,1:ExpandableCollapsibleArea,0:CommunityContent,1:TopicNotInScope,1:FeedViewerBasic,1:ImageSprite,2:Header.2,2:HeaderFooterSprite,2:Header.MediaQueries,2:Banner.MediaQueries,3:megabladeMenu.1,3:MegabladeMenu.MediaQueries,3:MegabladeMenuSpriteCluster,0:Breadcrumbs,0:Breadcrumbs.MediaQueries,0:ResponsiveToc,0:ResponsiveToc.MediaQueries,1:NavSidebar,0:LibraryMemberFilter,4:StandardRating,2:Footer.2,5:LinkList,2:Footer.MediaQueries,0:BaseResponsive,6:MsdnResponsive,0:Tables.MediaQueries,7:SkinnyRatingResponsive,7:SkinnyRatingV2;/Areas/Library/Content:0,/Areas/Epx/Content/Css:1,/Areas/Epx/Themes/TechNet/Content:2,/Areas/Epx/Themes/Shared/Content:3,/Areas/Global/Content:4,/Areas/Epx/Themes/Base/Content:5,/Areas/Library/Themes/Msdn/Content:6,/Areas/Library/Themes/TechNet/Content:7&amp;v=9192817066EC5D087D15C766A0430C95">
    <style>
        #execHeader td:first-child  {{ width: 10% ;}}
        #execHeader td:nth-child(3) {{ width: 35% ;}}
    </style>
</head>

<body lang=EN-US link=blue>
<div id="documentWrapper" style="width: 90%; margin-left: auto; margin-right: auto;">

<h1>Microsoft Security Bulletin Summary for {0}</h1>

<p>This bulletin summary lists security bulletins released for {0}.</p>

<p>Microsoft also provides information to help customers prioritize
monthly security updates with any non-security, high-priority updates that are being
released on the same day as the monthly security updates. Please see the section,
<b>Other Information</b>.
</p>

<p>
As a reminder, the <a href="https://portal.msrc.microsoft.com/en-us/security-guidance">Security Updates Guide</a> 
will be replacing security bulletins. Please see our blog post, 
<a href="https://blogs.technet.microsoft.com/msrc/2016/11/08/furthering-our-commitment-to-security-updates/">Furthering our commitment to security updates</a>, for more details.
</p>

<p>To receive automatic notifications whenever Microsoft Security
Updates are issued, subscribe to <a href="http://go.microsoft.com/fwlink/?LinkId=21163">Microsoft Technical Security Notifications</a>.
</p>

<h1>Executive Summaries</h1>

<p>The following table summarizes the security bulletins for this month in order of severity.
For details on affected software, see the next section, Affected Software.
</p>

<table border=1 cellpadding=0 width="99%">
 <thead style="background-color: #ededed">
  <tr>
   <td><b>CVE ID</b></td>
   <td><b>Vulnerability Description</b></td>
   <td><b>Maximum Severity Rating</b></td>
   <td><b>Vulnerability Impact</b></td>
   <td><b>Restart Requirement</b></td>
   <td><b>Known Issues</b></td>
   <td><b>Affected Software</b></td>
  </tr>
 </thead>
 {1}
</table>

<h1>Exploitability Index</h1>

<p>The following table provides an exploitability assessment of each of the vulnerabilities addressed this month. The vulnerabilities are listed in order of bulletin ID then CVE ID. Only vulnerabilities that have a severity rating of Critical or Important in the bulletins are included.</p>

<p><b>How do I use this table?</b></p>

<p>Use this table to learn about the likelihood of code execution and denial of service exploits within 30 days of security bulletin release, for each of the security updates that you may need to install. Review each of the assessments below, in accordance with your specific configuration, to prioritize your deployment of this month's updates. For more information about what these ratings mean, and how they are determined, please see <a href="http://technet.microsoft.com/security/cc998259">Microsoft Exploitability Index</a>.
</p>

<p>In the columns below, "Latest Software Release" refers to the subject software, and "Older Software Releases" refers to all older, supported releases of the subject software, as listed in the "Affected Software" and "Non-Affected Software" tables in the bulletin.</p>

<table border=1 cellpadding=0 width="99%">
 <thead style="background-color: #ededed">
  <tr>
   <td><b>CVE ID</b></td>
   <td><b>Vulnerability Title</b></td>
   <td><b>Exploitability Assessment for Latest Software Release</b></td>
   <td><b>Exploitability Assessment for Older Software Release</b></td>
   <td><b>Denial of Service Exploitability Assessment</b></td>   
  </tr>
 </thead>
 {2}
</table>

<h1>Affected Software</h1>

<p>The following tables list the bulletins in order of major software category and severity.</p>
<p>Use these tables to learn about the security updates that you may need to install. You should review each software program or component listed to see whether any security updates pertain to your installation. If a software program or component is listed, then the severity rating of the software update is also listed.</p>
<p><b>Note:</b> You may have to install several security updates for a single vulnerability. Review the whole column for each bulletin identifier that is listed to verify the updates that you have to install, based on the programs or components that you have installed on your system.</p>

<table border=1 cellpadding=0 width="99%">
 <thead style="background-color: #ededed">
  <tr>
   <td><b>Product</b></td>
   <td><b>KB Article</b></td>
   <td><b>Details</b></td>
   <td><b>Severity</b></td>  
   <td><b>Impact</b></td>  
  </tr>
 </thead>
{3}
</table>

<h1>Detection and Deployment Tools and Guidance</h1>

<p>Several resources are available to help administrators deploy security updates.</p>
<ul>
    <li>
        Microsoft Baseline Security Analyzer (MBSA) lets
        administrators scan local and remote systems for missing security updates and common
        security misconfigurations.
    </li>
    <li>
        Windows Server Update Services (WSUS), Systems Management Server (SMS), 
        and System Center Configuration Manager help administrators distribute security updates.
    </li>
    <li>
        The Update Compatibility Evaluator components included with Application Compatibility 
        Toolkit aid in streamlining the testing and validation of Windows updates against installed applications.
    </li>
</ul>

<p>For information about these and other tools that are available, see 
    <a href="http://technet.microsoft.com/security/cc297183">Security Tools for IT Pros</a>.
</p>

<h1>Other Information</h1>

<h2>Microsoft Windows Malicious Software Removal Tool</h2>

<p>Microsoft will release an updated version of the Microsoft Windows
Malicious Software Removal Tool on Windows Update, Microsoft Update, Windows Server
Update Services, and the Download Center.</p>

<h2>Microsoft Active Protections Program (MAPP)</h2>

<p>To improve security protections for customers, Microsoft provides
vulnerability information to major security software providers in advance of each
monthly security update release. Security software providers can then use this vulnerability
information to provide updated protections to customers via their security software
or devices, such as antivirus, network-based intrusion detection systems, or host-based
intrusion prevention systems. To determine whether active protections are available
from security software providers, please visit the active protections websites provided
by program partners, listed in 
<a href="http://go.microsoft.com/fwlink/?LinkId=215201">Microsoft Active Protections Program (MAPP) Partners</a>.
</p>

<h2>Security Strategies and Community</h2>

<p>Updates for other security issues are available from the following locations:</p>

<ul>
<li>
    Non-Window Security updates are available from <a href="http://go.microsoft.com/fwlink/?LinkId=21129">Microsoft Download Center</a>.
    You can find them most easily by doing a keyword search for &quot;security update&quot;.
</li>
<li>
    All Updates are available from <a href="http://go.microsoft.com/fwlink/?LinkID=40747">Microsoft Update</a>.
</li>
</ul>

<h2>IT Pro Security Community</h2>

<p>Learn to improve security and optimize your IT infrastructure,
and participate with other IT Pros on security topics in 
<a href="http://go.microsoft.com/fwlink/?LinkId=21164">IT Pro Security Community</a>.
</p>

<h2>Support</h2>
<ul>
<li>
    The affected software listed has been tested to determine
    which versions are affected. Other versions are past their support life cycle. To
    determine the support life cycle for your software version, visit 
    <a href="http://go.microsoft.com/fwlink/?LinkId=21742">Microsoft Support Lifecycle</a>.
</li>
<li>
    Help protect your computer that is running Windows
    from viruses and malware: 
    <a href="http://support.microsoft.com/contactus/cu_sc_virsec_master">Virus and Security Solution Center</a>
</li>
</ul>

<h2>Disclaimer</h2>

<p>The information provided in the Microsoft Knowledge Base is
provided &quot;as is&quot; without warranty of any kind. Microsoft disclaims all
warranties, either express or implied, including the warranties of merchantability
and fitness for a particular purpose. In no event shall Microsoft Corporation or
its suppliers be liable for any damages whatsoever including direct, indirect, incidental,
consequential, loss of business profits or special damages, even if Microsoft Corporation
or its suppliers have been advised of the possibility of such damages. Some states
do not allow the exclusion or limitation of liability for consequential or incidental
damages so the foregoing limitation may not apply.</p>

</div>

 </body>
</html>
'@ 

    #region CVE Summary Table
    $cveSummaryRowTemplate = @'
<tr>
     <td>{0}</td>
     <td>{1}</td>
     <td>{2}</td>
     <td>{3}</td>
     <td>{4}</td>
     <td>{5}</td>
     <td>{6}</td>
 </tr>
'@
    $cveSummaryTableHtml = ''

    foreach($vuln in $Vulnerability)
    {
        $SeverityValues = $vuln.Threats | Where-Object Type -EQ 3 | 
          Select @{Name='Severity' ;Expression={$_.Description.Value}} -Unique |
          Select -ExpandProperty Severity

        if ($SeverityValues -contains 'Critical')
        {
            $maximumSeverity = 'Critical'
        }
        elseif ($SeverityValues -contains 'Important')
        {
            $maximumSeverity = 'Important'
        }
        elseif ($SeverityValues -contains 'Moderate')
        {
            $maximumSeverity = 'Moderate'
        }
        elseif ($SeverityValues -contains 'Low')
        {
            $maximumSeverity = 'Low'
        }
        else
        {
            Write-Warning "Could not determine the Maximum Severity from the Threats"
            $maximumSeverity = 'Unknown'
        } 
        
        $ImpactValues = $vuln.Threats | Where-Object Type -EQ 0 | ForEach-Object {$_.Description.Value} | Select-Object -Unique   

        $AffectedSoftware = $vuln.ProductStatuses.ProductID | 
        ForEach-Object {
            $ProductTree.FullProductName | 
            Where ProductID -EQ $PSItem | 
            Select -ExpandProperty Value
        } | Select -Unique


        $cveSummaryTableHtml += $cveSummaryRowTemplate -f @(
            $vuln.CVE #TODO - make this an href
            $vuln.Notes | Where Title -eq Description | Select -ExpandProperty Value
            $maximumSeverity
            $ImpactValues -join ',<br>'
            'TODO'
            'TODO'
            $AffectedSoftware -join ',<br>'
        )
    }

    #endregion

    #region Exploitability Index Table
    $exploitabilityRowTemplate = @'
<tr>
     <td>{0}</td>
     <td>{1}</td>
     <td>{2}</td>
     <td>{3}</td>
     <td>{4}</td>
 </tr>
'@

    $exploitabilityIndexTableHtml = ''

    foreach($vuln in $Vulnerability)
    {
        $ExploitStatusLatest = ''
        $ExploitStatusOlder  = ''

        $ExploitStatusThreat = $vuln.Threats | Where Type -EQ 1 | Select -Last 1
        $ExploitStatus = Get-MsrcThreatExploitStatus -ExploitStatusString $ExploitStatusThreat.Description.Value        

        $exploitabilityIndexTableHtml += $exploitabilityRowTemplate -f @(
            $vuln.CVE #TODO - make this an href
            $vuln.Title.Value
            $ExploitStatus.LatestSoftwareRelease
            $ExploitStatus.OlderSoftwareRelease
            $ExploitStatus.DenialOfService           
        )
    }
    
    #endregion

    #region Affected Software Table
    
    $affectedSoftwareRowTemplate = @'
    <tr>
         <td>{0}</td>
         <td>{1}</td>
         <td>{2}</td>
         <td>{3}</td>
         <td>{4}</td>
    </tr>
'@

    $affectedSoftwareTableHtml = ''
    $affectedSoftware = Get-MsrcCvrfAffectedSoftware -Vulnerability $Vulnerability -ProductTree $ProductTree

    foreach($affectedSoftwareItem in $affectedSoftware)
    {        
        $affectedSoftwareTableHtml += $affectedSoftwareRowTemplate -f @(
            $affectedSoftwareItem.FullProductName
            $affectedSoftwareItem.KBArticle
            $affectedSoftwareItem.CVE
            $affectedSoftwareItem.Severity
            $affectedSoftwareItem.Impact
        )
    }

    #endregion

    Write-Output ($htmlDocumentTemplate -f @(
        $DocumentTitle.Value          #Title
        $cveSummaryTableHtml          #CVE Summary Rows
        $exploitabilityIndexTableHtml #Expoitability Rows
        $affectedSoftwareTableHtml    #Affected Software Rows
    ))

}

function Get-MsrcThreatExploitStatus
{
<#

.EXAMPLE
"Publicly Disclosed:No;Exploited:No;Latest Software Release:Exploitation More Likely;Older Software Release:N/A;" | Get-MsrcThreatExploitStatus
.EXAMPLE
Get-MsrcThreatExploitStatus -ExploitStatusString "Publicly Disclosed:No;Exploited:No;Latest Software Release:Exploitation More Likely;Older Software Release:N/A;"
#>
    Param
    (
        <#
        The Exploit Status string, which is delimited        
        #> 
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true)]
        $ExploitStatusString
    )
    Process
     {
        $ExploitStatus = [PSCustomObject]@{
            PubliclyDisclosed     = ''
            Exploited             = ''
            LatestSoftwareRelease = ''
            OlderSoftwareRelease  = ''
            DenialOfService       = 'N/A'
        }
        foreach ($exploitStatusItem in $ExploitStatusString -split ';')
        {
            $exploitStatusName, $exploitStatusValue = $exploitStatusItem -split ':'
            if ($exploitStatusName -eq 'Publicly Disclosed')
            {
                $ExploitStatus.PubliclyDisclosed = $exploitStatusValue
            }
            if ($exploitStatusName -eq 'Exploited')
            {
                $ExploitStatus.Exploited = $exploitStatusValue
            }
            if ($exploitStatusName -eq 'Latest Software Release')
            {
                $ExploitStatus.LatestSoftwareRelease = $exploitStatusValue
            }
            if ($exploitStatusName -eq 'Older Software Release')
            {
                $ExploitStatus.OlderSoftwareRelease = $exploitStatusValue
            }
        }
        Write-Output $ExploitStatus
    }
}

function Get-MsrcCvrfAffectedSoftware
{
<#
.Synopsis
   Get details of products affected by a CVRF document
.DESCRIPTION
   CVRF documents next products into several places, including:
   -Vulnerabilities
   -Threats
   -Remediations
   -Product Tree
   This function gathers the details for each product identified in a CVRF document.

.EXAMPLE
   #Get product details from a CVRF document using the pipeline   
   Get-MsrcCvrfDocument -ID 2016-Nov -ApiKey 'YOUR API KEY' | Get-MsrcCvrfAffectedSoftware
.EXAMPLE
   #Get product details from a CVRF document using a variable and parameters
   $cvrfDocument = Get-MsrcCvrfDocument -ID 2016-Nov -ApiKey 'YOUR API KEY'
   Get-MsrcCvrfAffectedSoftware -Vulnerability $cvrfDocument.Vulnerability -ProductTree $cvrfDocument.ProductTree
#>
    Param
    (
        <#
        API Key for the MSRC CVRF API
        To get an API key, visit https://portal.msrc.microsoft.com
        #> 
        [Parameter(Mandatory=$true, 
                   ValueFromPipelineByPropertyName=$true)]
        $Vulnerability,

        <#
        Get the CVRF document for the specified CVRF ID (ie. 2016-Aug)
        #>
        [Parameter(Mandatory=$true, 
                   ValueFromPipelineByPropertyName=$true)]
        $ProductTree
    )    

    foreach($vuln in $Vulnerability)
    {
        foreach($vulnProductID in $vuln.ProductStatuses.ProductID)
        {
    
            $FullProductName = $ProductTree.FullProductName | Where-Object ProductID -EQ $vulnProductID | Select -ExpandProperty Value        

            $KBArticle = $vuln.Remediations | Where-Object ProductID -Contains $vulnProductID | Select -ExpandProperty Description | Select -ExpandProperty Value
        
            $Severity = $vuln.Threats | Where Type -EQ 3 | Where-Object ProductID -Contains $vulnProductID | Select -ExpandProperty Description | Select -ExpandProperty Value

            $Impact = $vuln.Threats | Where Type -EQ 0 | Where-Object ProductID -Contains $vulnProductID | Select -ExpandProperty Description | Select -ExpandProperty Value

            [PSCustomObject] @{
                FullProductName = $FullProductName
                KBArticle       = $KBArticle
                CVE             = $vuln.CVE
                Severity        = $Severity
                Impact          = $Impact
            }
        }
    }
}

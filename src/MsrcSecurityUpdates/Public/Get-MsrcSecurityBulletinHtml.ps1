Function Get-MsrcSecurityBulletinHtml {
[CmdletBinding()]
Param(

    [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
    $Vulnerability,

    [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
    $ProductTree,


    [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
    $DocumentTracking,

    [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
    $DocumentTitle
)
Begin {
    $css = @'
        body {
          background-color: darkgray;
        }

        h1 {
          color: maroon;
        }
        table {
          font-family: Arial, Helvetica, sans-serif;
          border-collapse: collapse;
          width: 100%;
        }

        table td, th {
          border: 1px solid #ddd;
          padding: 8px;
        }

        table tr:nth-child(even){
          background-color: #ddd;
        }

        table tr:hover {background-color: #FAF0E6;}

        table th {
          padding-top: 12px;
          padding-bottom: 12px;
          text-align: left;
          background-color: #C0C0C0;
        }
'@
    $htmlDocumentTemplate = @'
<html>
<head>
    <!-- Created by module version {4} -->
    <!-- this is the css from the old bulletin site. Change this to better style your report to your liking -->
    <!-- <link rel="stylesheet" href="https://i-technet.sec.s-msft.com/Combined.css?resources=0:ImageSprite,0:TopicResponsive,0:TopicResponsive.MediaQueries,1:CodeSnippet,1:ProgrammingSelector,1:ExpandableCollapsibleArea,0:CommunityContent,1:TopicNotInScope,1:FeedViewerBasic,1:ImageSprite,2:Header.2,2:HeaderFooterSprite,2:Header.MediaQueries,2:Banner.MediaQueries,3:megabladeMenu.1,3:MegabladeMenu.MediaQueries,3:MegabladeMenuSpriteCluster,0:Breadcrumbs,0:Breadcrumbs.MediaQueries,0:ResponsiveToc,0:ResponsiveToc.MediaQueries,1:NavSidebar,0:LibraryMemberFilter,4:StandardRating,2:Footer.2,5:LinkList,2:Footer.MediaQueries,0:BaseResponsive,6:MsdnResponsive,0:Tables.MediaQueries,7:SkinnyRatingResponsive,7:SkinnyRatingV2;/Areas/Library/Content:0,/Areas/Epx/Content/Css:1,/Areas/Epx/Themes/TechNet/Content:2,/Areas/Epx/Themes/Shared/Content:3,/Areas/Global/Content:4,/Areas/Epx/Themes/Base/Content:5,/Areas/Library/Themes/Msdn/Content:6,/Areas/Library/Themes/TechNet/Content:7&amp;v=9192817066EC5D087D15C766A0430C95"> -->

    <!-- this style section changes cell widths in the exec header table so that the affected products at the end are wide enough to read -->
    <style>
{5}
        #execHeader td:first-child  {{ width: 10% ;}}
        #execHeader td:nth-child(5) {{ width: 37% ;}}
    </style>

    <!-- this section defines explicit width for all cells in the affected software tables. This is so the column width is the same across each product -->
    <style>
        .affected_software td:first-child {{ width: 20% ; }}
        .affected_software td:nth-child(2) {{ width: 20% ; }}
        .affected_software td:nth-child(3) {{ width: 15% ; }}
        .affected_software td:nth-child(4) {{ width: 22.5% ; }}
        .affected_software td:nth-child(5) {{ width: 22.5% ; }}

    </style>

</head>

<body lang=EN-US link=blue>
<div id="documentWrapper" style="width: 90%; margin-left: auto; margin-right: auto;">

<h1>Microsoft Security Bulletin Summary for {0}</h1>

<p>This document is a summary for Microsoft security updates released for {0}.</p>

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

<p>The following table summarizes the security updates for this month in order of severity.
For details on affected software, see the next section, Affected Software.
</p>

<table id="execHeader" border=1 cellpadding=0 width="99%">
 <thead style="background-color: #ededed">
  <tr>
   <td><b>CVE ID</b></td>
   <td><b>Vulnerability Description</b></td>
   <td><b>Maximum Severity Rating</b></td>
   <td><b>Vulnerability Impact</b></td>
   <td><b>Affected Software</b></td>
  </tr>
 </thead>
 {1}
</table>

<h1>Exploitability Index</h1>

<p>The following table provides an <a href="https://www.microsoft.com/en-us/msrc/exploitability-index?rtc=1">exploitability assessment</a> for this vulnerability at the time of original publication.</p>

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

<!-- Affected software tables -->
{3}
<!-- End Affected software tables -->

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

    $cveSummaryRowTemplate = @'
<tr>
     <td>{0}</td>
     <td>{1}</td>
     <td>{2}</td>
     <td>{3}</td>
     <td>{4}</td>
 </tr>
'@
    $cveSummaryTableHtml = ''

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

    $affectedSoftwareNameHeaderTemplate = @'
    <table class="affected_software" border=1 cellpadding=0 width="99%">
        <thead style="background-color: #ededed">
            <tr>
                <td colspan="5"><b>{0}</b></td>
            </tr>
        </thead>
            <tr>
                <td><b>CVE ID</b></td>
                <td><b>KB Article</b></td>
                <td><b>Restart Required</b></td>
                <td><b>Severity</b></td>
                <td><b>Impact</b></td>
            </tr>
        {1}
    </table>
'@

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
    $affectedSoftwareDocumentHtml = ''
}
Process {

    #region CVE Summary Table
    $HT = @{
        Vulnerability = $PSBoundParameters['Vulnerability']
        ProductTree = $PSBoundParameters['ProductTree']
    }

    Get-MsrcCvrfCVESummary @HT |
    ForEach-Object {
        $cveSummaryTableHtml += $cveSummaryRowTemplate -f @(
            "$($_.CVE)<br><a href=`"https://cve.mitre.org/cgi-bin/cvename.cgi?name=$($_.CVE)`">MITRE</a><br><a href=`"https://web.nvd.nist.gov/view/vuln/detail?vulnId=$($_.CVE)`">NVD</a>"
            $_.Description
            $_.'Maximum Severity Rating'
            $_.'Vulnerability Impact' -join ',<br>'
            $_.'Affected Software' -join ',<br>'
        )
    }
    #endregion

    #region Exploitability Index Table

    Get-MsrcCvrfExploitabilityIndex -Vulnerability $PSBoundParameters['Vulnerability'] |
    ForEach-Object {
        $exploitabilityIndexTableHtml += $exploitabilityRowTemplate -f @(
            $_.CVE #TODO - make this an href
            $_.Title
            $_.LatestSoftwareRelease
            $_.OlderSoftwareRelease
            'N/A' # was $ExploitStatus.DenialOfService
        )
    }
    #endregion

    #region Affected Software Table

    $affectedSoftware = Get-MsrcCvrfAffectedSoftware @HT

    $affectedSoftware.FullProductName |
    Sort-Object -Unique |
    ForEach-Object {

        $PN = $_

        $affectedSoftwareTableHtml = ''

        $affectedSoftware |
        Where-Object { $_.FullProductName -eq $PN } |
        Sort-Object -Unique -Property CVE |
        ForEach-Object {
            $affectedSoftwareTableHtml += $affectedSoftwareRowTemplate -f @(
                $_.CVE,
                $(
                    if (-not($_.KBArticle)) {
                        'None'
                    } else {
                        ($_.KBArticle | ForEach-Object {
                            '<a href="{0}">{1}</a><br>' -f  $_.URL, $_.ID
                        }) -join '<br />'
                    }
                ),
                $(
                    if (-not($_.RestartRequired)) {
                        'Unknown'
                    } else{
                        ($_.RestartRequired | ForEach-Object {
                            '{0}<br>' -f $_
                        })  -join '<br />'
                    }
                ),
                $(
                    if (-not($_.Severity)) {
                        'Unknown'
                    } else {
                        ($_.Severity | ForEach-Object {
                            '{0}<br>' -f $_
                        })  -join '<br />'
                    }
                ),
                $(
                    if (-not($_.Impact)) {
                        'Unknown'
                    } else {
                        ($_.Impact | ForEach-Object {
                            '{0}<br>' -f $_
                        }) -join '<br />'
                    }
                )
            )
        }
        $affectedSoftwareDocumentHtml += $affectedSoftwareNameHeaderTemplate -f @(
            $PN
            $affectedSoftwareTableHtml
        )
    }
    #endregion

    $htmlDocumentTemplate -f @(
        $DocumentTitle.Value,           # Title
        $cveSummaryTableHtml,           # CVE Summary Rows
        $exploitabilityIndexTableHtml,  # Expoitability Rows
        $affectedSoftwareDocumentHtml,  # Affected Software Rows
        "$($MyInvocation.MyCommand.Version.ToString())",
        $css
    )

}
End {}
}
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

    $htmlDocumentTemplate = @'
<html>
<head>
    <!-- this is the css from the old bulletin site. Change this to better style your report to your liking -->
    <link rel="stylesheet" href="https://i-technet.sec.s-msft.com/Combined.css?resources=0:ImageSprite,0:TopicResponsive,0:TopicResponsive.MediaQueries,1:CodeSnippet,1:ProgrammingSelector,1:ExpandableCollapsibleArea,0:CommunityContent,1:TopicNotInScope,1:FeedViewerBasic,1:ImageSprite,2:Header.2,2:HeaderFooterSprite,2:Header.MediaQueries,2:Banner.MediaQueries,3:megabladeMenu.1,3:MegabladeMenu.MediaQueries,3:MegabladeMenuSpriteCluster,0:Breadcrumbs,0:Breadcrumbs.MediaQueries,0:ResponsiveToc,0:ResponsiveToc.MediaQueries,1:NavSidebar,0:LibraryMemberFilter,4:StandardRating,2:Footer.2,5:LinkList,2:Footer.MediaQueries,0:BaseResponsive,6:MsdnResponsive,0:Tables.MediaQueries,7:SkinnyRatingResponsive,7:SkinnyRatingV2;/Areas/Library/Content:0,/Areas/Epx/Content/Css:1,/Areas/Epx/Themes/TechNet/Content:2,/Areas/Epx/Themes/Shared/Content:3,/Areas/Global/Content:4,/Areas/Epx/Themes/Base/Content:5,/Areas/Library/Themes/Msdn/Content:6,/Areas/Library/Themes/TechNet/Content:7&amp;v=9192817066EC5D087D15C766A0430C95">
    
    <!-- this style section changes cell widths in the exec header table so that the affected products at the end are wide enough to read -->
    <style>
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
            "$($_.CVE)<br><a href=`"http://www.cve.mitre.org/cgi-bin/cvename.cgi?name=$($_.CVE)`">MITRE</a><br><a href=`"https://web.nvd.nist.gov/view/vuln/detail?vulnId=$($_.CVE)`">NVD</a>"
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
        ForEach-Object {
            $affectedSoftwareTableHtml += $affectedSoftwareRowTemplate -f @(
                $_.CVE,
                $(
                    if (-not($_.KBArticle)) {
                        'None'
                    } else {
                        $_.KBArticle | ForEach-Object {
                            '<a href="https://catalog.update.microsoft.com/v7/site/Search.aspx?q={0}">{0}</a><br>' -f  $_
                        }
                    }
                ),
                $(
                    if (-not($_.RestartRequired)) {
                        'Unknown'
                    } else{
                        $_.RestartRequired | ForEach-Object {
                            '{0}<br>' -f $_
                        }
                    }
                ),
                $(
                    if (-not($_.Severity)) {
                        'Unknown'
                    } else {
                        $_.Severity | ForEach-Object {
                            '{0}<br>' -f $_
                        }
                    }
                ),
                $(
                    if (-not($_.Impact)) {
                        'Unknown'
                    } else { 
                        $_.Impact | ForEach-Object {
                            '{0}<br>' -f $_
                        }
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
        $DocumentTitle.Value           # Title
        $cveSummaryTableHtml           # CVE Summary Rows
        $exploitabilityIndexTableHtml  # Expoitability Rows
        $affectedSoftwareDocumentHtml  # Affected Software Rows
    )

}
End {}
}
# SIG # Begin signature block
# MIIkYQYJKoZIhvcNAQcCoIIkUjCCJE4CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAAmFNvWA0fWDg3
# Jam/lh4ZolwJxoLDDDPQn01lI8S99aCCDZMwggYRMIID+aADAgECAhMzAAAAjoeR
# pFcaX8o+AAAAAACOMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMTYxMTE3MjIwOTIxWhcNMTgwMjE3MjIwOTIxWjCBgzEL
# MAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1v
# bmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjENMAsGA1UECxMETU9Q
# UjEeMBwGA1UEAxMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMIIBIjANBgkqhkiG9w0B
# AQEFAAOCAQ8AMIIBCgKCAQEA0IfUQit+ndnGetSiw+MVktJTnZUXyVI2+lS/qxCv
# 6cnnzCZTw8Jzv23WAOUA3OlqZzQw9hYXtAGllXyLuaQs5os7efYjDHmP81LfQAEc
# wsYDnetZz3Pp2HE5m/DOJVkt0slbCu9+1jIOXXQSBOyeBFOmawJn+E1Zi3fgKyHg
# 78CkRRLPA3sDxjnD1CLcVVx3Qv+csuVVZ2i6LXZqf2ZTR9VHCsw43o17lxl9gtAm
# +KWO5aHwXmQQ5PnrJ8by4AjQDfJnwNjyL/uJ2hX5rg8+AJcH0Qs+cNR3q3J4QZgH
# uBfMorFf7L3zUGej15Tw0otVj1OmlZPmsmbPyTdo5GPHzwIDAQABo4IBgDCCAXww
# HwYDVR0lBBgwFgYKKwYBBAGCN0wIAQYIKwYBBQUHAwMwHQYDVR0OBBYEFKvI1u2y
# FdKqjvHM7Ww490VK0Iq7MFIGA1UdEQRLMEmkRzBFMQ0wCwYDVQQLEwRNT1BSMTQw
# MgYDVQQFEysyMzAwMTIrYjA1MGM2ZTctNzY0MS00NDFmLWJjNGEtNDM0ODFlNDE1
# ZDA4MB8GA1UdIwQYMBaAFEhuZOVQBdOCqhc3NyK1bajKdQKVMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY0Nv
# ZFNpZ1BDQTIwMTFfMjAxMS0wNy0wOC5jcmwwYQYIKwYBBQUHAQEEVTBTMFEGCCsG
# AQUFBzAChkVodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NlcnRzL01p
# Y0NvZFNpZ1BDQTIwMTFfMjAxMS0wNy0wOC5jcnQwDAYDVR0TAQH/BAIwADANBgkq
# hkiG9w0BAQsFAAOCAgEARIkCrGlT88S2u9SMYFPnymyoSWlmvqWaQZk62J3SVwJR
# avq/m5bbpiZ9CVbo3O0ldXqlR1KoHksWU/PuD5rDBJUpwYKEpFYx/KCKkZW1v1rO
# qQEfZEah5srx13R7v5IIUV58MwJeUTub5dguXwJMCZwaQ9px7eTZ56LadCwXreUM
# tRj1VAnUvhxzzSB7pPrI29jbOq76kMWjvZVlrkYtVylY1pLwbNpj8Y8zon44dl7d
# 8zXtrJo7YoHQThl8SHywC484zC281TllqZXBA+KSybmr0lcKqtxSCy5WJ6PimJdX
# jrypWW4kko6C4glzgtk1g8yff9EEjoi44pqDWLDUmuYx+pRHjn2m4k5589jTajMW
# UHDxQruYCen/zJVVWwi/klKoCMTx6PH/QNf5mjad/bqQhdJVPlCtRh/vJQy4njpI
# BGPveJiiXQMNAtjcIKvmVrXe7xZmw9dVgh5PgnjJnlQaEGC3F6tAE5GusBnBmjOd
# 7jJyzWXMT0aYLQ9RYB58+/7b6Ad5B/ehMzj+CZrbj3u2Or2FhrjMvH0BMLd7Hald
# G73MTRf3bkcz1UDfasouUbi1uc/DBNM75ePpEIzrp7repC4zaikvFErqHsEiODUF
# he/CBAANa8HYlhRIFa9+UrC4YMRStUqCt4UqAEkqJoMnWkHevdVmSbwLnHhwCbww
# ggd6MIIFYqADAgECAgphDpDSAAAAAAADMA0GCSqGSIb3DQEBCwUAMIGIMQswCQYD
# VQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEe
# MBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMTIwMAYDVQQDEylNaWNyb3Nv
# ZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkgMjAxMTAeFw0xMTA3MDgyMDU5
# MDlaFw0yNjA3MDgyMTA5MDlaMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNo
# aW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29y
# cG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBIDIw
# MTEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCr8PpyEBwurdhuqoIQ
# TTS68rZYIZ9CGypr6VpQqrgGOBoESbp/wwwe3TdrxhLYC/A4wpkGsMg51QEUMULT
# iQ15ZId+lGAkbK+eSZzpaF7S35tTsgosw6/ZqSuuegmv15ZZymAaBelmdugyUiYS
# L+erCFDPs0S3XdjELgN1q2jzy23zOlyhFvRGuuA4ZKxuZDV4pqBjDy3TQJP4494H
# DdVceaVJKecNvqATd76UPe/74ytaEB9NViiienLgEjq3SV7Y7e1DkYPZe7J7hhvZ
# PrGMXeiJT4Qa8qEvWeSQOy2uM1jFtz7+MtOzAz2xsq+SOH7SnYAs9U5WkSE1JcM5
# bmR/U7qcD60ZI4TL9LoDho33X/DQUr+MlIe8wCF0JV8YKLbMJyg4JZg5SjbPfLGS
# rhwjp6lm7GEfauEoSZ1fiOIlXdMhSz5SxLVXPyQD8NF6Wy/VI+NwXQ9RRnez+ADh
# vKwCgl/bwBWzvRvUVUvnOaEP6SNJvBi4RHxF5MHDcnrgcuck379GmcXvwhxX24ON
# 7E1JMKerjt/sW5+v/N2wZuLBl4F77dbtS+dJKacTKKanfWeA5opieF+yL4TXV5xc
# v3coKPHtbcMojyyPQDdPweGFRInECUzF1KVDL3SV9274eCBYLBNdYJWaPk8zhNqw
# iBfenk70lrC8RqBsmNLg1oiMCwIDAQABo4IB7TCCAekwEAYJKwYBBAGCNxUBBAMC
# AQAwHQYDVR0OBBYEFEhuZOVQBdOCqhc3NyK1bajKdQKVMBkGCSsGAQQBgjcUAgQM
# HgoAUwB1AGIAQwBBMAsGA1UdDwQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB8GA1Ud
# IwQYMBaAFHItOgIxkEO5FAVO4eqnxzHRI4k0MFoGA1UdHwRTMFEwT6BNoEuGSWh0
# dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY1Jvb0Nl
# ckF1dDIwMTFfMjAxMV8wM18yMi5jcmwwXgYIKwYBBQUHAQEEUjBQME4GCCsGAQUF
# BzAChkJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0Nl
# ckF1dDIwMTFfMjAxMV8wM18yMi5jcnQwgZ8GA1UdIASBlzCBlDCBkQYJKwYBBAGC
# Ny4DMIGDMD8GCCsGAQUFBwIBFjNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtp
# b3BzL2RvY3MvcHJpbWFyeWNwcy5odG0wQAYIKwYBBQUHAgIwNB4yIB0ATABlAGcA
# YQBsAF8AcABvAGwAaQBjAHkAXwBzAHQAYQB0AGUAbQBlAG4AdAAuIB0wDQYJKoZI
# hvcNAQELBQADggIBAGfyhqWY4FR5Gi7T2HRnIpsLlhHhY5KZQpZ90nkMkMFlXy4s
# PvjDctFtg/6+P+gKyju/R6mj82nbY78iNaWXXWWEkH2LRlBV2AySfNIaSxzzPEKL
# UtCw/WvjPgcuKZvmPRul1LUdd5Q54ulkyUQ9eHoj8xN9ppB0g430yyYCRirCihC7
# pKkFDJvtaPpoLpWgKj8qa1hJYx8JaW5amJbkg/TAj/NGK978O9C9Ne9uJa7lryft
# 0N3zDq+ZKJeYTQ49C/IIidYfwzIY4vDFLc5bnrRJOQrGCsLGra7lstnbFYhRRVg4
# MnEnGn+x9Cf43iw6IGmYslmJaG5vp7d0w0AFBqYBKig+gj8TTWYLwLNN9eGPfxxv
# FX1Fp3blQCplo8NdUmKGwx1jNpeG39rz+PIWoZon4c2ll9DuXWNB41sHnIc+BncG
# 0QaxdR8UvmFhtfDcxhsEvt9Bxw4o7t5lL+yX9qFcltgA1qFGvVnzl6UJS0gQmYAf
# 0AApxbGbpT9Fdx41xtKiop96eiL6SJUfq/tHI4D1nvi/a7dLl+LrdXga7Oo3mXkY
# S//WsyNodeav+vyL6wuA6mk7r/ww7QRMjt/fdW1jkT3RnVZOT7+AVyKheBEyIXrv
# QQqxP/uozKRdwaGIm1dxVk5IRcBCyZt2WwqASGv9eZ/BvW1taslScxMNelDNMYIW
# JDCCFiACAQEwgZUwfjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEoMCYGA1UEAxMfTWljcm9zb2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMQITMwAA
# AI6HkaRXGl/KPgAAAAAAjjANBglghkgBZQMEAgEFAKCCAREwGQYJKoZIhvcNAQkD
# MQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJ
# KoZIhvcNAQkEMSIEIFB0twD7oq24wznJbyhqTtlFCBsBcbQ00icj5tC/9E/kMIGk
# BgorBgEEAYI3AgEMMYGVMIGSoEyASgBNAHMAcgBjAFMAZQBjAHUAcgBpAHQAeQBV
# AHAAZABhAHQAZQBzACAAUABvAHcAZQByAFMAaABlAGwAbAAgAE0AbwBkAHUAbABl
# oUKAQGh0dHBzOi8vZ2l0aHViLmNvbS9NaWNyb3NvZnQvTVNSQy1NaWNyb3NvZnQt
# U2VjdXJpdHktVXBkYXRlcy1BUEkwDQYJKoZIhvcNAQEBBQAEggEAjL8nngtxgQA4
# zCYkOgp8Vk1i3lccbHXwdUmQmzOr0yNiTdx14RGsMUxmPUD0yySsXo/SEeedMMDm
# y8NvmkamUDCTjZcaNGtbglcQXT3sJwznMGsFQyn3n3LJw3Yy1v6Ygtn99dZ6kxDo
# Dq4a9pI72SdjlP8h/5JWlauAtNSeYJOUOP/jfCKwKaQlJgniUogWBYirPvDMPDVr
# 0wJgwwahYqfW2p77V8c3m+UmqY7AblFScbPfouFeRWP2GQSuQLnv7jgO0ZgXsWB6
# XYf7Hj7JutspuebUPkv+7rnfL/mpjX5vZhzxkaRfEjgfqe7bB267/mQlArpL20Cu
# gtZTKb7CCKGCE0owghNGBgorBgEEAYI3AwMBMYITNjCCEzIGCSqGSIb3DQEHAqCC
# EyMwghMfAgEDMQ8wDQYJYIZIAWUDBAIBBQAwggE9BgsqhkiG9w0BCRABBKCCASwE
# ggEoMIIBJAIBAQYKKwYBBAGEWQoDATAxMA0GCWCGSAFlAwQCAQUABCCgXiqR5bDZ
# 7upMjcPjV1ZPCCIsVqU6F+L9iv0GVJpL+wIGWNVNFYsNGBMyMDE3MDQxMzE4MDgy
# OC4wMDVaMAcCAQGAAgH0oIG5pIG2MIGzMQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMQ0wCwYDVQQLEwRNT1BSMScwJQYDVQQLEx5uQ2lwaGVyIERT
# RSBFU046QkJFQy0zMENBLTJEQkUxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0
# YW1wIFNlcnZpY2Wggg7NMIIGcTCCBFmgAwIBAgIKYQmBKgAAAAAAAjANBgkqhkiG
# 9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAO
# BgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEy
# MDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5IDIw
# MTAwHhcNMTAwNzAxMjEzNjU1WhcNMjUwNzAxMjE0NjU1WjB8MQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGlt
# ZS1TdGFtcCBQQ0EgMjAxMDCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEB
# AKkdDbx3EYo6IOz8E5f1+n9plGt0VBDVpQoAgoX77XxoSyxfxcPlYcJ2tz5mK1vw
# FVMnBDEfQRsalR3OCROOfGEwWbEwRA/xYIiEVEMM1024OAizQt2TrNZzMFcmgqNF
# DdDq9UeBzb8kYDJYYEbyWEeGMoQedGFnkV+BVLHPk0ySwcSmXdFhE24oxhr5hoC7
# 32H8RsEnHSRnEnIaIYqvS2SJUGKxXf13Hz3wV3WsvYpCTUBR0Q+cBj5nf/VmwAOW
# RH7v0Ev9buWayrGo8noqCjHw2k4GkbaICDXoeByw6ZnNPOcvRLqn9NxkvaQBwSAJ
# k3jN/LzAyURdXhacAQVPIk0CAwEAAaOCAeYwggHiMBAGCSsGAQQBgjcVAQQDAgEA
# MB0GA1UdDgQWBBTVYzpcijGQ80N7fEYbxTNoWoVtVTAZBgkrBgEEAYI3FAIEDB4K
# AFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSME
# GDAWgBTV9lbLj+iiXGJo0T2UkFvXzpoYxDBWBgNVHR8ETzBNMEugSaBHhkVodHRw
# Oi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJB
# dXRfMjAxMC0wNi0yMy5jcmwwWgYIKwYBBQUHAQEETjBMMEoGCCsGAQUFBzAChj5o
# dHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dF8y
# MDEwLTA2LTIzLmNydDCBoAYDVR0gAQH/BIGVMIGSMIGPBgkrBgEEAYI3LgMwgYEw
# PQYIKwYBBQUHAgEWMWh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9QS0kvZG9jcy9D
# UFMvZGVmYXVsdC5odG0wQAYIKwYBBQUHAgIwNB4yIB0ATABlAGcAYQBsAF8AUABv
# AGwAaQBjAHkAXwBTAHQAYQB0AGUAbQBlAG4AdAAuIB0wDQYJKoZIhvcNAQELBQAD
# ggIBAAfmiFEN4sbgmD+BcQM9naOhIW+z66bM9TG+zwXiqf76V20ZMLPCxWbJat/1
# 5/B4vceoniXj+bzta1RXCCtRgkQS+7lTjMz0YBKKdsxAQEGb3FwX/1z5Xhc1mCRW
# S3TvQhDIr79/xn/yN31aPxzymXlKkVIArzgPF/UveYFl2am1a+THzvbKegBvSzBE
# JCI8z+0DpZaPWSm8tv0E4XCfMkon/VWvL/625Y4zu2JfmttXQOnxzplmkIz/amJ/
# 3cVKC5Em4jnsGUpxY517IW3DnKOiPPp/fZZqkHimbdLhnPkd/DjYlPTGpQqWhqS9
# nhquBEKDuLWAmyI4ILUl5WTs9/S/fmNZJQ96LjlXdqJxqgaKD4kWumGnEcua2A5H
# moDF0M2n0O99g/DhO3EJ3110mCIIYdqwUB5vvfHhAN/nMQekkzr3ZUd46PioSKv3
# 3nJ+YWtvd6mBy6cJrDm77MbL2IK0cs0d9LiFAR6A+xuJKlQ5slvayA1VmXqHczsI
# 5pgt6o3gMy4SKfXAL1QnIffIrE7aKLixqduWsqdCosnPGUFN4Ib5KpqjEWYw07t0
# MkvfY3v1mYovG8chr1m1rtxEPJdQcdeh0sVV42neV8HR3jDA/czmTfsNv11P6Z0e
# GTgvvM9YBS7vDaBQNdrvCScc1bN+NR4Iuto229Nfj950iEkSMIIE2jCCA8KgAwIB
# AgITMwAAAKGl/bnup/yenQAAAAAAoTANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQg
# VGltZS1TdGFtcCBQQ0EgMjAxMDAeFw0xNjA5MDcxNzU2NDhaFw0xODA5MDcxNzU2
# NDhaMIGzMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMQ0wCwYD
# VQQLEwRNT1BSMScwJQYDVQQLEx5uQ2lwaGVyIERTRSBFU046QkJFQy0zMENBLTJE
# QkUxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0G
# CSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCb0AF55iKvenwbdR7zt2fWHd6n3gA0
# 2BL5B+oIex9NTcqFlHqrdmsqB0WNUSfTtl1GpSXGYhR5i0/M5iz17J9Neh00IjYq
# uMPi7MVQ7dP9Q8Etv8Xw+s/MhJDroqaybVegj7lhcNRJzogvgy47gUqTtUlKxGXJ
# loXkL/qs4thXHTP2vhDnwlIbE+D5FDaos5v02xXw9NJrfS24Vc4R6Vb/lOkhDruR
# V8ycFXlwzY6s0+OBmZjDDgff23PFzylj9T7sNxh6c/YkdbX8yTeUMFcH1aBAFU0L
# FrDm1TddPNjTq7yHl4d6VXNLYUPB8wmIkr7OuOWESjwWN5xBziCXgcgJAgMBAAGj
# ggEbMIIBFzAdBgNVHQ4EFgQUq1y5gr5xKtWP4OiyKc/6G9O3+kgwHwYDVR0jBBgw
# FoAU1WM6XIoxkPNDe3xGG8UzaFqFbVUwVgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDov
# L2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljVGltU3RhUENB
# XzIwMTAtMDctMDEuY3JsMFoGCCsGAQUFBwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0
# cDovL3d3dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNUaW1TdGFQQ0FfMjAx
# MC0wNy0wMS5jcnQwDAYDVR0TAQH/BAIwADATBgNVHSUEDDAKBggrBgEFBQcDCDAN
# BgkqhkiG9w0BAQsFAAOCAQEAI2BMuAVU6WfJvjk8FTLDE8izAJB/FzOb/XYvXUwS
# s+iwJKL+svQfqLKOLk4QB1zo9zSK7jQd6OFEhvge2949EpSHwZoPQ+Cb+hRhq7bq
# EuOqiGrXZXflB1vQUFPRxVrUKC8qSlF5H3k5KnfHYeUjfyoF2iae1UC24l/cOhN0
# 5Tr9qvs/Avwr+fggUlsoyl2yICjuHR70ioS8F1LqsxJxmiwdG04NeNHbkw0kXheI
# SVQh/NhcJtDpE+Fsyk6/B7g7+eGcL0YMZTqcRbAJp3NMLGu21xZj4PxyOJmmBc0y
# kUGiXvq7160Oe4XL8w93O3gy00+WkRRII8aKl5dYHf2lQaGCA3YwggJeAgEBMIHj
# oYG5pIG2MIGzMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMQ0w
# CwYDVQQLEwRNT1BSMScwJQYDVQQLEx5uQ2lwaGVyIERTRSBFU046QkJFQy0zMENB
# LTJEQkUxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2WiJQoB
# ATAJBgUrDgMCGgUAAxUAgq6J9bTmPxZcILwX5MHDQM0562iggcIwgb+kgbwwgbkx
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xDTALBgNVBAsTBE1P
# UFIxJzAlBgNVBAsTHm5DaXBoZXIgTlRTIEVTTjo1N0Y2LUMxRTAtNTU0QzErMCkG
# A1UEAxMiTWljcm9zb2Z0IFRpbWUgU291cmNlIE1hc3RlciBDbG9jazANBgkqhkiG
# 9w0BAQUFAAIFANyaLMkwIhgPMjAxNzA0MTMxNjU4NDlaGA8yMDE3MDQxNDE2NTg0
# OVowdDA6BgorBgEEAYRZCgQBMSwwKjAKAgUA3JosyQIBADAHAgEAAgIasDAHAgEA
# AgIYNzAKAgUA3Jt+SQIBADA2BgorBgEEAYRZCgQCMSgwJjAMBgorBgEEAYRZCgMB
# oAowCAIBAAIDFuNgoQowCAIBAAIDB6EgMA0GCSqGSIb3DQEBBQUAA4IBAQBOOBCy
# ZCrxc4dOV3tqtkqjDFUxQQDoSha5cwMvoCUqlC+0nP/nKVZ+sEuwb0Mc2YC3RO8c
# O3zF8gQ9xQuCU43MJcyWmvov4HDdS5fZMv1HMyRhQgC42+sCJkCZq74vMxqMVm2h
# HuMQanCoUKmNbS+p70qzbaP/n2Vn/+iMwGxaIAOBf2EBcE2EdFig7oN2JRUy1otw
# IyRD9u52wEaqZzQmKmqTPl5KNR1S5akerAxBN7weB5sguAgVNNzqm5hBNYqdZxOb
# Eo3F1Xn/aez6C0bnKzA2N6kf/CEjoN5KosmHEGCV/XVQhC9jQlt0ANu+CdZVpjiT
# KLdE3im7RTXQDGh6MYIC9TCCAvECAQEwgZMwfDELMAkGA1UEBhMCVVMxEzARBgNV
# BAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
# c29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAg
# UENBIDIwMTACEzMAAAChpf257qf8np0AAAAAAKEwDQYJYIZIAWUDBAIBBQCgggEy
# MBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAvBgkqhkiG9w0BCQQxIgQgE4en
# qlJRjadIhUAbRaWz9gCpxgeVZ9SHwyjVr4awEIMwgeIGCyqGSIb3DQEJEAIMMYHS
# MIHPMIHMMIGxBBSCron1tOY/FlwgvBfkwcNAzTnraDCBmDCBgKR+MHwxCzAJBgNV
# BAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4w
# HAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29m
# dCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMzAAAAoaX9ue6n/J6dAAAAAAChMBYEFDa4
# ngI4z4EmiPIcikXsYdV7rhoDMA0GCSqGSIb3DQEBCwUABIIBAAlA9w09fWUcsz8j
# +dUgy++JDA5vWJdpoAU45Dg+eMMgx1p7ZNR1+Zc0FljL1ChLGXWijCPMZ01EDRgp
# bF2DAKBLmQF2AlUSRSBxdn8pA0Cv8chkBB55liWDCq+BF8RPpHMEiptO/Y2RnkTe
# 9xQ2txsHAyxNNnpR8yTCmw/5EMZSdr8070Yb8TYsQuUSxK0ioQBdapAFL0ka/v/E
# yUNypYiumry0OsJ2Srg85udDRCE2dZGyaElIpqmXkYrg/oRucwEN0BxxbstTxHXU
# C1/Q7FlO8P1aLwgBsH2MNQT36yWEgrsTawPl8qvI6bFrQyF2vWsridKUYg48/9Jr
# o2YIHgg=
# SIG # End signature block

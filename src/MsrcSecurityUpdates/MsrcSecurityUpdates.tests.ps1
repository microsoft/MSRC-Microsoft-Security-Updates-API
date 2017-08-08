
# Import module would only work if the module is found in standard locations
# Import-Module -Name MsrcSecurityUpdates -Force
Import-Module .\MsrcSecurityUpdates.psd1 -Verbose -Force

# Set-MSRCApiKey -ApiKey "API KEY PLACEHOLDER" -Verbose

if (-not ($global:MSRCApiKey)) {

   Write-Warning -Message 'You need to use Set-MSRCApiKey first to set your API Key'
   break
}

<#
Get-Help Get-MsrcSecurityUpdate
Get-Help Get-MsrcSecurityUpdate -Examples

Get-Help Get-MsrcCvrfDocument
Get-Help Get-MsrcCvrfDocument -Examples

Get-Help Get-MsrcSecurityBulletinHtml
Get-Help Get-MsrcSecurityBulletinHtml -Examples

Get-Help Get-MsrcCvrfAffectedSoftware
Get-Help Get-MsrcCvrfAffectedSoftware -Examples
#> 

Describe 'Function: Get-MsrcSecurityUpdateMSRC (calls the /Updates API)' {

    It 'Get-MsrcSecurityUpdate - all' {
        Get-MsrcSecurityUpdate | 
        Should Not BeNullOrEmpty 
    }

    It 'Get-MsrcSecurityUpdate - by year' {
        Get-MsrcSecurityUpdate -Year 2017 | 
        Should Not BeNullOrEmpty 
    }

    It 'Get-MsrcSecurityUpdate - by vulnerability' {
        Get-MsrcSecurityUpdate -Vulnerability CVE-2017-0003 | 
        Should Not BeNullOrEmpty 
    }

    It 'Get-MsrcSecurityUpdate - by cvrf' {
        Get-MsrcSecurityUpdate -Cvrf 2017-Jan | 
        Should Not BeNullOrEmpty 
    }

    It 'Get-MsrcSecurityUpdate - by date - before' {
        Get-MsrcSecurityUpdate -Before 2017-01-01 | 
        Should Not BeNullOrEmpty 
    }

    It 'Get-MsrcSecurityUpdate - by date - after' {
        Get-MsrcSecurityUpdate -After 2017-01-01 | 
        Should Not BeNullOrEmpty 
    }

    It 'Get-MsrcSecurityUpdate - by date - before and after' {
        Get-MsrcSecurityUpdate -Before 2017-01-01 -After 2016-10-01 | 
        Should Not BeNullOrEmpty 
    }
}

Describe 'Function: Get-MsrcCvrfDocument (calls the MSRC /cvrf API)' {

    It 'Get-MsrcCvrfDocument - 2016-Nov' {
        Get-MsrcCvrfDocument -ID 2016-Nov | 
        Should Not BeNullOrEmpty 
    }

    It 'Get-MsrcCvrfDocument - 2016-Nov - as XML' {
        Get-MsrcCvrfDocument -ID 2016-Nov -AsXml | 
        Should Not BeNullOrEmpty 
    }

    Get-MsrcSecurityUpdate | 
    Foreach-Object {
        It 'Get-MsrcCvrfDocument - none shall throw: $($PSItem.ID)' {
            {
                Get-MsrcCvrfDocument -ID $PSItem.ID | 
                Out-Null
            } |
            Should Not Throw
        }
    }
}

Describe 'Function: Set-MSRCApiKey with proxy' {
    if (-not ($global:msrcProxy)) {

       Write-Warning -Message 'This test requires you to use Set-MSRCApiKey first to set your API Key and proxy details'
       break
    }

    It 'Get-MsrcSecurityUpdate - all' {
        Get-MsrcSecurityUpdate | 
        Should Not BeNullOrEmpty 
    }

    It 'Get-MsrcCvrfDocument - 2016-Nov' {
        Get-MsrcCvrfDocument -ID 2016-Nov | 
        Should Not BeNullOrEmpty 
    }
}

# May still work but not ready yet...
# Describe 'Function: Get-MsrcSecurityBulletinHtml (generates the MSRC Security Bulletin HTML Report)' {
#     It 'Security Bulletin Report' {
#         Get-MsrcCvrfDocument -ID 2016-Nov |
#         Get-MsrcSecurityBulletinHtml |
#         Should Not BeNullOrEmpty
#     }
# }
InModuleScope MsrcSecurityUpdates {
    Describe 'Function: Get-MsrcCvrfAffectedSoftware' {
        It 'Get-MsrcCvrfAffectedSoftware by pipeline' {
            Get-MsrcCvrfDocument -ID 2016-Nov |
            Get-MsrcCvrfAffectedSoftware |
            Should Not BeNullOrEmpty
        }

        It 'Get-MsrcCvrfAffectedSoftware by parameters' {
            $cvrfDocument = Get-MsrcCvrfDocument -ID 2016-Nov
            Get-MsrcCvrfAffectedSoftware -Vulnerability $cvrfDocument.Vulnerability -ProductTree $cvrfDocument.ProductTree |
            Should Not BeNullOrEmpty
        }
    }

    Describe 'Function: Get-MsrcCvrfProductVulnerability' {
        It 'Get-MsrcCvrfProductVulnerability by pipeline' {
            Get-MsrcCvrfDocument -ID 2016-Nov |
            Get-MsrcCvrfProductVulnerability |
            Should Not BeNullOrEmpty
        }

        It 'Get-MsrcCvrfProductVulnerability by parameters' {
            $cvrfDocument = Get-MsrcCvrfDocument -ID 2016-Nov
            Get-MsrcCvrfProductVulnerability -Vulnerability $cvrfDocument.Vulnerability -ProductTree $cvrfDocument.ProductTree -DocumentTracking $cvrfDocument.DocumentTracking -DocumentTitle $cvrfDocument.DocumentTitle  |
            Should Not BeNullOrEmpty
        }
    }
}

Describe 'Function: Get-MsrcVulnerabilityReportHtml (generates the MSRC Vulnerability Summary HTML Report)' {
    It 'Vulnerability Summary Report - does not throw' {
        {
            Get-MsrcCvrfDocument -ID 2016-Nov |
            Get-MsrcVulnerabilityReportHtml -Verbose | Out-Null
        } |
        Should Not Throw
    }

    #Get-MsrcSecurityUpdate | 
    #Foreach-Object {
    #    It 'Vulnerability Summary Report - none shall throw: $($PSItem.ID)' {
    #        {
    #            Get-MsrcCvrfDocument -ID $PSItem.ID |
    #            Get-MsrcVulnerabilityReportHtml | 
    #            Out-Null
    #        } |
    #        Should Not Throw
    #    }
    #}
}

InModuleScope MsrcSecurityUpdates {
	Describe 'Function: Get-KBDownloadUrl (generates the html for KBArticle downloads used in the vulnerability report affected software table)' {
		It 'Get-KBDownloadUrl by pipeline' {
			{
				$doc = Get-MsrcCvrfDocument -ID 2017-May
				$af = $doc | Get-MsrcCvrfAffectedSoftware 
				$af.KBArticle | Get-KBDownloadUrl
			} |
			Should Not Throw
		}


		It 'Get-KBDownloadUrl by parameters' {
			{
				$doc = Get-MsrcCvrfDocument -ID 2017-May
				$af = $doc | Get-MsrcCvrfAffectedSoftware 
				Get-KBDownloadUrl -KBArticleObject $af.KBArticle
			} |
			Should Not Throw
		}
	}
}

#When a pester test fails, it writes out to stdout, and sets an error in $Error. When invoking powershell from C# it is a lot easier to read the stderr stream.
if($Error)
{
    Write-Error 'A pester test has failed during the validation process'
}

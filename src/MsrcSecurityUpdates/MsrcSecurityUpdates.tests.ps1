
# Import module would only work if the module is found in standard locations
# Import-Module -Name MsrcSecurityUpdates -Force
$Error.Clear()
Get-Module -Name MsrcSecurityUpdates | Remove-Module -Force -Verbose:$false
Import-Module (Join-Path -Path $PSScriptRoot -ChildPath 'MsrcSecurityUpdates.psd1') -Verbose:$false -Force

Describe 'API version after module loading' {
    It '$msrcApiUrl = https://api.msrc.microsoft.com/cvrf/v3.0' {
        $msrcApiUrl -eq 'https://api.msrc.microsoft.com/cvrf/v3.0' | Should Be $true
    }
    It '$msrcApiVersion = api-version=2023-11-01' {
        $msrcApiVersion -eq 'api-version=2023-11-01' | Should Be $true
    }
    Set-MSRCApiKey -APIVersion 2.0
    It '$msrcApiUrl = https://api.msrc.microsoft.com/cvrf/v2.0' {
        $msrcApiUrl -eq 'https://api.msrc.microsoft.com/cvrf/v2.0' | Should Be $true
    }
    It '$msrcApiVersion = api-version=2016-08-01' {
        $msrcApiVersion -eq 'api-version=2016-08-01' | Should Be $true
    }
}

'2.0','3.0' |
Foreach-Object {
    $v = $_
    Set-MSRCApiKey -APIVersion $_
Describe ('Function: Get-MsrcSecurityUpdate (calls the /Updates API version {0})' -f $v) {

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
        { $w = $null
        $w = Get-MsrcSecurityUpdate -Before 2018-01-01 -WarningVariable w -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
        $w.Message -eq 'No results returned from the /Update API' } | Should Be $true
    }

    It 'Get-MsrcSecurityUpdate - by date - after' {
        { $w = $null
        Get-MsrcSecurityUpdate -After 2017-01-01 -WarningVariable w -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
        $w.Message -eq 'No results returned from the /Update API' } | Should Be $true
    }

    It 'Get-MsrcSecurityUpdate - by date - before and after' {
        { $w = $null
        Get-MsrcSecurityUpdate -Before 2018-01-01 -After 2017-10-01 -WarningVariable w -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
        $w.Message -eq 'No results returned from the /Update API' } | Should Be $true
    }
}

Describe ('Function: Get-MsrcCvrfDocument (calls the MSRC /cvrf API version {0})' -f $v) {

    It 'Get-MsrcCvrfDocument - 2016-Nov' {
        Get-MsrcCvrfDocument -ID 2016-Nov |
        Should Not BeNullOrEmpty
    }

    It 'Get-MsrcCvrfDocument - 2016-Nov - as XML' {
        Get-MsrcCvrfDocument -ID 2016-Nov -AsXml |
        Should Not BeNullOrEmpty
    }

<#
    Get-MsrcSecurityUpdate | Where-Object { $_.ID -ne '2017-May-B' } |
    Where-Object { $_.ID -eq "$(((Get-Date).AddMonths(-1)).ToString('yyyy-MMM',[System.Globalization.CultureInfo]'en-US'))" } |
    Foreach-Object {
        It "Get-MsrcCvrfDocument - none shall throw: $($PSItem.ID)" {
            {
                $null = Get-MsrcCvrfDocument -ID $PSItem.ID
            } |
            Should Not Throw
        }
    }

    It 'Get-MsrcCvrfDocument for 2017-May-B with Get-MsrcCvrfDocument should throw' {
        {
            Get-MsrcSecurityUpdate | Where-Object { $_.ID -eq '2017-May-B' } |
            Foreach-Object {
                $null = Get-MsrcCvrfDocument -ID $PSItem.ID
            }
        } | Should Throw
    }
#>
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

Describe ('Function: Get-MsrcVulnerabilityReportHtml (generates the MSRC Vulnerability Summary HTML Report with API version {0})' -f $v) {
    It 'Vulnerability Summary Report - does not throw' {
        {
            $null = Get-MsrcCvrfDocument -ID 2016-Nov |
            Get-MsrcVulnerabilityReportHtml -Verbose:$false -ShowNoProgress -WarningAction SilentlyContinue
        } |
        Should Not Throw
    }
<#
    Get-MsrcSecurityUpdate | Where-Object { $_.ID -ne '2017-May-B' } |
    Where-Object { $_.ID -eq "$(((Get-Date).AddMonths(-1)).ToString('yyyy-MMM',[System.Globalization.CultureInfo]'en-US'))" } |
    Foreach-Object {
        It "Vulnerability Summary Report - none shall throw: $($PSItem.ID)" {
            {
                $null = Get-MsrcCvrfDocument -ID $PSItem.ID |
                Get-MsrcVulnerabilityReportHtml -ShowNoProgress -WarningAction SilentlyContinue
            } |
            Should Not Throw
        }
    }
#>
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

}

#When a pester test fails, it writes out to stdout, and sets an error in $Error. When invoking powershell from C# it is a lot easier to read the stderr stream.
if($Error)
{
    Write-Error -Message 'A pester test has failed during the validation process'
}

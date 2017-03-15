

$msrcApiKey = ''

Import-Module -Name MsrcSecurityUpdates -Force

Get-Help Get-MsrcSecurityUpdate
Get-Help Get-MsrcSecurityUpdate -Examples

Get-Help Get-MsrcCvrfDocument
Get-Help Get-MsrcCvrfDocument -Examples

Get-Help Get-MsrcSecurityBulletinHtml
Get-Help Get-MsrcSecurityBulletinHtml -Examples

Get-Help Get-MsrcCvrfAffectedSoftware
Get-Help Get-MsrcCvrfAffectedSoftware -Examples

Describe 'Function: Get-MsrcSecurityUpdateMSRC (calls the /Updates API)' {

    It "Get-MsrcSecurityUpdate - all" {
        Get-MsrcSecurityUpdate -ApiKey $msrcApiKey -Verbose | 
        Should Not BeNullOrEmpty 
    }

    It "Get-MsrcSecurityUpdate - by year" {
        Get-MsrcSecurityUpdate -ApiKey $msrcApiKey -Year 2017 -Verbose | 
        Should Not BeNullOrEmpty 
    }

    It "Get-MsrcSecurityUpdate - by vulnerability" {
        Get-MsrcSecurityUpdate -ApiKey $msrcApiKey -Vulnerability CVE-2017-0003 -Verbose | 
        Should Not BeNullOrEmpty 
    }

    It "Get-MsrcSecurityUpdate - by cvrf" {
        Get-MsrcSecurityUpdate -ApiKey $msrcApiKey -Cvrf 2017-Jan -Verbose | 
        Should Not BeNullOrEmpty 
    }

    It "Get-MsrcSecurityUpdate - by date - before" {
        Get-MsrcSecurityUpdate -ApiKey $msrcApiKey -Before 2017-01-01 -Verbose | 
        Should Not BeNullOrEmpty 
    }

    It "Get-MsrcSecurityUpdate - by date - after" {
        Get-MsrcSecurityUpdate -ApiKey $msrcApiKey -After 2017-01-01 -Verbose | 
        Should Not BeNullOrEmpty 
    }

    It "Get-MsrcSecurityUpdate - by date - before and after" {
        Get-MsrcSecurityUpdate -ApiKey $msrcApiKey -Before 2017-01-01 -After 2016-10-01 -Verbose | 
        Should Not BeNullOrEmpty 
    }
}

Describe 'Function: Get-MsrcCvrfDocument (calls the MSRC /cvrf API)' {

    It "Get-MsrcCvrfDocument - 2016-Nov" {
        Get-MsrcCvrfDocument -ID 2016-Nov -ApiKey $msrcApiKey -Verbose | 
        Should Not BeNullOrEmpty 
    }

    It "Get-MsrcCvrfDocument - 2016-Nov - as JSON" {
        Get-MsrcCvrfDocument -ID 2016-Nov -ApiKey $msrcApiKey -AsJson -Verbose | 
        Should Not BeNullOrEmpty 
    }

    It "Get-MsrcCvrfDocument - 2016-Nov - as XML" {
        Get-MsrcCvrfDocument -ID 2016-Nov -ApiKey $msrcApiKey -AsXml -Verbose | 
        Should Not BeNullOrEmpty 
    }
}

Describe 'Function: Get-MsrcSecurityBulletinHtml (generates the MSRC Security Bulletin HTML Report)' {
    It 'Security Bulletin Report' {
        Get-MsrcCvrfDocument -ID 2016-Nov -ApiKey $msrcApiKey -Verbose |
        Get-MsrcSecurityBulletinHtml -Verbose |
        Should Not BeNullOrEmpty
    }
}

Describe 'Function: Get-MsrcCvrfAffectedSoftware' {
    It 'Get-MsrcCvrfAffectedSoftware by pipeline' {
        Get-MsrcCvrfDocument -ID 2016-Nov -ApiKey $msrcApiKey -Verbose |
        Get-MsrcCvrfAffectedSoftware -Verbose |
        Should Not BeNullOrEmpty
    }

    It 'Get-MsrcCvrfAffectedSoftware by parameters' {
        $cvrfDocument = Get-MsrcCvrfDocument -ID 2016-Nov -ApiKey $msrcApiKey -Verbose
        Get-MsrcCvrfAffectedSoftware -Vulnerability $cvrfDocument.Vulnerability -ProductTree $cvrfDocument.ProductTree |
        Should Not BeNullOrEmpty
    }
}

Describe 'Function: Get-MsrcCvrfProductVulnerability' {
    It 'Get-MsrcCvrfProductVulnerability by pipeline' {
        Get-MsrcCvrfDocument -ID 2016-Nov -ApiKey $msrcApiKey -Verbose |
        Get-MsrcCvrfProductVulnerability -Verbose |
        Should Not BeNullOrEmpty
    }

    It 'Get-MsrcCvrfProductVulnerability by parameters' {
        $cvrfDocument = Get-MsrcCvrfDocument -ID 2016-Nov -ApiKey $msrcApiKey -Verbose
        Get-MsrcCvrfProductVulnerability -Vulnerability $cvrfDocument.Vulnerability -ProductTree $cvrfDocument.ProductTree -DocumentTracking $cvrfDocument.DocumentTracking -DocumentTitle $cvrfDocument.DocumentTitle  |
        Should Not BeNullOrEmpty
    }
}
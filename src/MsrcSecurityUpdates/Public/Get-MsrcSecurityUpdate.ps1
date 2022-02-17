Function Get-MsrcSecurityUpdate {
<#
    .SYNOPSIS
       Get MSRC security updates

    .DESCRIPTION
       Calls the CVRF Update API to get a list of security updates

    .PARAMETER After
        Get security updates released after this date

    .PARAMETER Before
        Get security updates released before this date

    .PARAMETER Year
        Get security updates for the specified year (ie. 2016)

    .PARAMETER Vulnerability
        Get security updates for the specified Vulnerability CVE (ie. CVE-2016-0128)

    .PARAMETER Cvrf
        Get security update for the specified CVRF ID (ie. 2016-Aug)

    .EXAMPLE
       Get-MsrcSecurityUpdate

       Get all the updates

    .EXAMPLE
       Get-MsrcSecurityUpdate -Vulnerability CVE-2017-0003

       Get all the updates containing Vulnerability CVE-2017-003

    .EXAMPLE
       Get-MsrcSecurityUpdate -Year 2017

       Get all the updates for the year 2017

    .EXAMPLE
       Get-MsrcSecurityUpdate -Cvrf 2017-Jan

       Get all the updates for the CVRF document with ID of 2017-Jan

    .EXAMPLE
       Get-MsrcSecurityUpdate -Before 2017-01-01

       Get all the updates before January 1st, 2017

    .EXAMPLE
       Get-MsrcSecurityUpdate -After 2017-01-01

       Get all the updates after January 1st, 2017

    .EXAMPLE
       Get-MsrcSecurityUpdate -Before 2017-01-01 -After 2016-10-01

       Get all the updates before January 1st, 2017 and after October 1st, 2016

    .EXAMPLE
        Get-MsrcSecurityUpdate -After (Get-Date).AddDays(-60) -Before (Get-Date)

        Get all updates between now and the last 60 days

    .NOTES
        An API Key for the MSRC CVRF API is required
        To get an API key, please visit https://portal.msrc.microsoft.com

#>
[CmdletBinding(DefaultParameterSetName='All')]
Param (

    [Parameter(ParameterSetName='ByDate')]
    [DateTime]$After,

    [Parameter(ParameterSetName='ByDate')]
    [DateTime]$Before,

    [Parameter(Mandatory,ParameterSetName='ByYear')]
    [ValidateScript({
        if ($_ -lt 2016 -or $_ -gt [DateTime]::Now.Year) {
            throw 'Year must be between 2016 and this year'
        } else {
            $true
        }
    })]
    [Int]$Year,

    [Parameter(Mandatory,ParameterSetName='ByVulnerability')]
    [String]$Vulnerability
)
DynamicParam {
    $Dictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary

    $ParameterName = 'CVRF'
    $AttribColl1 = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
    $Param1Att = New-Object System.Management.Automation.ParameterAttribute
    $Param1Att.Mandatory = $true
    # $Param1Att.ValueFromPipeline = $true
    $Param1Att.ParameterSetName = 'ByCVRF'
    $AttribColl1.Add($Param1Att)

    try {
        $allCVRFID = Get-CVRFID
    } catch {
        Throw "`nUnable to get online the list of CVRF ID because:`n$($_.Exception.Message)"
    }
    if ($allCVRFID) {
        $AttribColl1.Add((New-Object System.Management.Automation.ValidateSetAttribute($allCVRFID)))
        $Dictionary.Add($ParameterName,(New-Object System.Management.Automation.RuntimeDefinedParameter($ParameterName, [string], $AttribColl1)))

        $Dictionary
    }
}
Begin {}
Process {

    switch ($PSCmdlet.ParameterSetName) {

        ByDate {

            $sb = New-Object System.Text.StringBuilder

            $null = $sb.Append("$($msrcApiUrl)/Updates?`$filter=")

            if ($PSBoundParameters.ContainsKey('Before')) {

                $null = $sb.Append("CurrentReleaseDate lt $($Before.ToString('yyyy-MM-dd'))")

                if ($PSBoundParameters.ContainsKey('After')) {
                    $null = $sb.Append(' and ')
                }

            }

            if ($PSBoundParameters.ContainsKey('After')) {

                $null = $sb.Append("CurrentReleaseDate gt $($After.ToString('yyyy-MM-dd'))")

            }

            $null = $sb.Append("&$($msrcApiVersion)")

            $url = $sb.ToString()

            break
        }
        ByYear {
            $url = "{0}/Updates('{1}')?{2}" -f $msrcApiUrl,$Year,$msrcApiVersion
            break
        }
        ByVulnerability {
            $url = "{0}/Updates('{1}')?{2}" -f $msrcApiUrl,$Vulnerability,$msrcApiVersion
            break
        }
        ByCVRF {
            $url = "{0}/Updates('{1}')?{2}" -f $msrcApiUrl,$($PSBoundParameters['CVRF']),$msrcApiVersion
            break
        }
        Default {
            $url = "{0}/Updates?{1}" -f $msrcApiUrl,$msrcApiVersion
        }
    }

    $RestMethod = @{
        uri = $url
        Headers = @{ 'Accept' = 'application/json' }
        ErrorAction = 'Stop'
    }
    if ($global:msrcProxy){
        $RestMethod.Add('Proxy' , $global:msrcProxy)
    }
    if ($global:msrcProxyCredential){
        $RestMethod.Add('ProxyCredential' , $global:msrcProxyCredential)
    }
    if ($global:MSRCAdalAccessToken)
    {
        $RestMethod.Headers.Add('Authorization' , $global:MSRCAdalAccessToken.CreateAuthorizationHeader())
    }

    try {

        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Write-Verbose -Message "Calling $($RestMethod.uri)"

        $r = Invoke-RestMethod @RestMethod

    } catch {
        Write-Error -Message "HTTP Get failed with status code $($_.Exception.Response.StatusCode): $($_.Exception.Response.StatusDescription)"
    }

    if (-not $r) {
        Write-Warning -Message 'No results returned from the /Update API'
    } else {
        $r.Value
    }

}
End {}
}
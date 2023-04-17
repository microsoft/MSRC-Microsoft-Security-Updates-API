Function Get-MsrcCvrfDocument {
<#
    .SYNOPSIS
        Get a MSRC CVRF document

    .DESCRIPTION
       Calls the MSRC CVRF API to get a CVRF document by ID

    .PARAMETER ID
        Get the CVRF document for the specified CVRF ID (ie. 2016-Aug)

    .PARAMETER AsXml
        Get the output as Xml

    .EXAMPLE
       Get-MsrcCvrfDocument -ID 2016-Aug

       Get the Cvrf document '2016-Aug' (returns an object converted from the CVRF JSON)

    .EXAMPLE
       Get-MsrcCvrfDocument -ID 2016-Aug -AsXml

       Get the Cvrf document '2016-Aug' (returns an object converted from CVRF XML)

    .NOTES
        An API Key for the MSRC CVRF API is required
        To get an API key, please visit https://portal.msrc.microsoft.com

#>
[CmdletBinding()]
Param (
    [Parameter(ParameterSetName='XmlOutput')]
    [Switch]$AsXml

)
DynamicParam {

    $Dictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary

    $ParameterName = 'ID'
    $AttribColl1 = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
    $Param1Att = New-Object System.Management.Automation.ParameterAttribute
    $Param1Att.Mandatory = $true
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

    # Common
    $RestMethod = @{
        uri = '{0}/cvrf/{1}?{2}' -f $msrcApiUrl,$PSBoundParameters['ID'],$msrcApiVersion
        Headers = @{
            'Accept' = if($AsXml){'application/xml'} else {'application/json'}
        }
        ErrorAction = 'Stop'
    }

    # Add proxy and creds if required
    if ($global:msrcProxy) {

        $RestMethod.Add('Proxy', $global:msrcProxy)

    }

    if ($global:msrcProxyCredential) {

        $RestMethod.Add('ProxyCredential',$global:msrcProxyCredential)

    }

    if ($global:MSRCAdalAccessToken) {

        $RestMethod.Headers.Add('Authorization', $global:MSRCAdalAccessToken.CreateAuthorizationHeader())

    }

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Write-Verbose -Message "Calling $($RestMethod.uri)"

        Invoke-RestMethod @RestMethod

    } catch {
        Write-Error -Message "HTTP Get failed with status code $($_.Exception.Response.StatusCode): $($_.Exception.Response.StatusDescription)"
    }

}
End {}
}
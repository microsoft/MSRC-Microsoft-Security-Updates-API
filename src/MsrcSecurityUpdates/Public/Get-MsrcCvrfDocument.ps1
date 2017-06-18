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

    if (-not ($global:MSRCApiKey -or $global:MSRCAdalAccessToken)) {

	    Write-Warning -Message 'You need to use Set-MSRCApiKey first to set your API Key'

    } else {  
        $Dictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary

        $ParameterName = 'ID'
        $AttribColl1 = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $Param1Att = New-Object System.Management.Automation.ParameterAttribute
        $Param1Att.Mandatory = $true        
        $AttribColl1.Add($Param1Att)

        try {
            $allCVRFID = Get-CVRFID
        } catch {
            Throw 'Unable to get online the list of CVRF ID'
        }
        if ($allCVRFID) {
            $AttribColl1.Add((New-Object System.Management.Automation.ValidateSetAttribute($allCVRFID)))
            $Dictionary.Add($ParameterName,(New-Object System.Management.Automation.RuntimeDefinedParameter($ParameterName, [string], $AttribColl1)))
        
            $Dictionary
        }
    }
}
Begin {}
Process {

    # Common
    $RestMethod = @{
        uri = '{0}/cvrf/{1}?{2}' -f $msrcApiUrl,$PSBoundParameters['ID'],$msrcApiVersion
        ErrorAction = 'Stop'
    }
    
    # Add proxy and creds if required
    if ($global:msrcProxy) {
        
        $RestMethod.Add('Proxy', $global:msrcProxy)
    
    }
    if ($global:msrcProxyCredential) {
        
        $RestMethod.Add('ProxyCredential',$global:msrcProxyCredential)
    
    }

    # Adjust header based on our variables
    if ($global:MSRCApiKey) {
        
        $RestMethod.Add('Header',@{ 'Api-Key' = $global:MSRCApiKey })
    
    } elseif ($global:MSRCAdalAccessToken) {
        
        $RestMethod.Add('Header',@{ 'Authorization' = $global:MSRCAdalAccessToken.CreateAuthorizationHeader() })
    
    } else {
        
        Write-Warning -Message 'You need to use Set-MSRCApiKey first to set your API Key'        
    
    }

    # If we have a header defined, we proceed
    if ($RestMethod['Header']) {
        
        if ($AsXml) {
            $RestMethod.Header.Add('Accept','application/xml')
        } else {
            $RestMethod.Header.Add('Accept','application/json')
        }

        try {
    
            Write-Verbose -Message "Calling $($RestMethod.uri)"

            $response = Invoke-RestMethod @RestMethod
     
        } catch {
            Write-Error "HTTP Get failed with status code $($_.Exception.Response.StatusCode): $($_.Exception.Response.StatusDescription)"       
        }
    
        # Invoke-RestMethod will return an string on PowerShell 4.0 and earlier
        # if the JSON-formatted response is larger than about two million characters
        if (-not $AsXml -and $response -is [string]) {
            $response = ParseJsonString($response)
        }
        $response
    }
}
End {}
}
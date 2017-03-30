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
    [Parameter(Mandatory)]
    [ValidateScript({
        if(Test-CVRFID -ID $_) {
            $true
        } else {
            Throw 'Sorry CVRF ID entered is not valid'
        }
    })]
    [String]$ID,

    [Parameter(ParameterSetName='XmlOutput')]
    [Switch]$AsXml

)
Begin {}
Process {

    if (-not ($global:MSRCApiKey)) {

	    Write-Warning -Message 'You need to use Set-MSRCApiKey first to set your API Key'

    } else {  
        $url = '{0}/cvrf/{1}?{2}' -f $msrcApiUrl,$ID,$msrcApiVersion

        $Header = @{ 'Api-Key' = $global:MSRCApiKey }

        if ($AsXml) {
            $Header.Add('Accept','application/xml')
        } else {
            $Header.Add('Accept','application/json')
        }
        try {
    
            Write-Verbose -Message "Calling $($url)"

            # Invoke-WebRequest -Uri $url -Headers $Header -ErrorAction Stop #).Content
            Invoke-RestMethod -Uri $url -Headers $Header -ErrorAction Stop
     
        } catch {
            Write-Error "HTTP Get failed with status code $($_.Exception.Response.StatusCode): $($_.Exception.Response.StatusDescription)"       
        }
    }
}
End {}
}
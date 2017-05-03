#Requires -Version 3.0

Function Set-MSRCApiKey {
[CmdletBinding(SupportsShouldProcess)]
Param(
	[Parameter(Mandatory)]
	$ApiKey,
    [Parameter()]
    [String]$Proxy,
    [Parameter()]
    [System.Management.Automation.PSCredential]$ProxyCredential
)
Begin {}
Process {
    if ($PSCmdlet.ShouldProcess($ApiKey,'Set item')) {

	    $global:MSRCApiKey = $ApiKey
        Write-Verbose -Message "Successfully set your API Key required by cmdlets of this module.  Calls to the MSRC APIs will now use your API key."

        # we also set other shared variables
        $global:msrcApiUrl     = 'https://api.msrc.microsoft.com'
        Write-Verbose -Message "Successfully defined a msrcApiUrl global variable that points to $($global:msrcApiUrl)"

        $global:msrcApiVersion = 'api-version=2016-08-01'
        Write-Verbose -Message "Successfully defined a msrcApiVersion global variable that points to $($global:msrcApiVersion)"

        if ($ProxyCredential){
            $global:msrcProxyCredential = $ProxyCredential
            Write-Verbose -Message 'Successfully defined a msrcProxyCredential global variable'
        }

        if ($Proxy) {
            $global:msrcProxy = $Proxy
            Write-Verbose -Message "Successfully defined a msrcProxyCredential global variable that points to $($global:msrcProxy)"
        }
        
        if ($global:MSRCAdalAccessToken)
        {
            Remove-Variable -Name MSRCAdalAccessToken -Scope Global
        }
    }
}
End {}
}
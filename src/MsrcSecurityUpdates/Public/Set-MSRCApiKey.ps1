Function Set-MSRCApiKey {
[CmdletBinding(SupportsShouldProcess)]
Param(
   [Parameter()]
    $ApiKey,

    [Parameter()]
    [System.Uri]$Proxy,

    [Parameter()]
    [ValidateNotNull()]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.Credential()]
    $ProxyCredential = [System.Management.Automation.PSCredential]::Empty,

    [Parameter()]
    [ValidateSet('3.0','2.0')]
    [System.String]$APIVersion='3.0'
)
Begin {}
Process {
    if ($PSCmdlet.ShouldProcess($ApiKey,'Set item')) {

        if ($ApiKey) {
                $global:MSRCApiKey = $ApiKey
                Write-Verbose -Message "Successfully set your API Key required by cmdlets of this module.  Calls to the MSRC APIs will now use your API key."
        }
        # we also set other shared variables
        $global:msrcApiUrl     = 'https://api.msrc.microsoft.com/cvrf/v{0}' -f $APIVersion
        Write-Verbose -Message "Successfully defined a msrcApiUrl global variable that points to $($global:msrcApiUrl)"

        $global:msrcApiVersion = Switch ($APIVersion) {
         '2.0' { 'api-version=2016-08-01'}
         '3.0' { 'api-version=2023-11-01'}
        }
        Write-Verbose -Message "Successfully defined a msrcApiVersion global variable that points to $($global:msrcApiVersion)"

        if ($ProxyCredential -ne [System.Management.Automation.PSCredential]::Empty) {
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
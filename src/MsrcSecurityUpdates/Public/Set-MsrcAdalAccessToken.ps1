Function Set-MSRCAdalAccessToken {
[CmdletBinding(SupportsShouldProcess)]
Param()
Begin {}
Process {
    if ($PSCmdlet.ShouldProcess('Set the MSRCApiKey using MSRCAdalAccessToken')) {
        $authority = 'https://login.windows.net/microsoft.onmicrosoft.com/'

        $authContext = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext($authority)

        $rUri = New-Object System.Uri -ArgumentList 'https://msrc-api-powershell'

        $promptBehavior = [Microsoft.IdentityModel.Clients.ActiveDirectory.PromptBehavior]::Auto

        $ResourceId = 'https://msrc-api-prod.azurewebsites.net'

        $ClientId = 'c7fe3b9e-4d97-462d-ae1b-c16e679be355'

        $authResult = $authContext.AcquireToken($ResourceId, $ClientId, $rUri,$promptBehavior)

	    $global:MSRCAdalAccessToken = $authResult
        Write-Verbose -Message "Successfully set your Access Token required by cmdlets of this module.    Calls to the MSRC APIs will now use your access token."

        # we also set other shared variables
        $global:msrcApiUrl     = 'https://api.msrc.microsoft.com'
        Write-Verbose -Message "Successfully defined a msrcApiUrl global variable that points to $($global:msrcApiUrl)"

        $global:msrcApiVersion = 'api-version=2016-08-01'
        Write-Verbose -Message "Successfully defined a msrcApiVersion global variable that points to $($global:msrcApiVersion)"

        if ($global:MSRCApiKey)
        {
            Remove-Variable -Name MSRCApiKey -Scope Global
        }
    }
}
End {}
}
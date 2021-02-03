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

        $platformParams = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters($promptBehavior)

        $ResourceId = 'https://msrc-api-prod.azurewebsites.net'

        $ClientId = 'c7fe3b9e-4d97-462d-ae1b-c16e679be355'

        $authResult = $authContext.AcquireTokenAsync($ResourceId, $ClientId, $rUri,$platformParams).GetAwaiter().GetResult()

	    $global:MSRCAdalAccessToken = $authResult
        Write-Verbose -Message "Successfully set your Access Token required by cmdlets of this module.    Calls to the MSRC APIs will now use your access token."
    }
}
End {}
}
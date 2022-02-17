Function Set-MSRCAdalAccessToken {
[CmdletBinding(SupportsShouldProcess)]
Param()
Begin {}
Process {
    if ([AppDomain]::CurrentDomain.SetupInformation.TargetFrameworkName -like "*v5.*") {
        throw ".Net Core v5.x is not currently supported"
    }

    if ($PSCmdlet.ShouldProcess('Set the MSRCApiKey using MSRCAdalAccessToken')) {
        Add-Type -Path "$PSScriptRoot/../Microsoft.IdentityModel.Clients.ActiveDirectory.dll" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue

        $authority = 'https://login.windows.net/microsoft.onmicrosoft.com/'

        $authContext = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext($authority)

        $rUri = New-Object System.Uri -ArgumentList 'https://msrc-api-powershell'

        $promptBehavior = [Microsoft.IdentityModel.Clients.ActiveDirectory.PromptBehavior]::Auto


        $ResourceId = 'https://msrc-api-prod.azurewebsites.net'

        $ClientId = 'c7fe3b9e-4d97-462d-ae1b-c16e679be355'

        $global:MSRCAdalAccessToken = $null

        if ($null -ne $authContext.AcquireToken) {
            $global:MSRCAdalAccessToken = $authContext.AcquireToken($ResourceId, $ClientId, $rUri,$promptBehavior)
        } elseif ($null -ne $authContext.AcquireTokenAsync) {
            $platformParams = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters($promptBehavior)
            $task = $authContext.AcquireTokenAsync($ResourceId, $ClientId, $rUri,$platformParams)
            $task.Wait()
            $global:MSRCAdalAccessToken = $task.Result
        }

	    if ($null -ne $global:MSRCAdalAccessToken) {
            Write-Verbose -Message "Successfully set your Access Token required by cmdlets of this module.    Calls to the MSRC APIs will now use your access token."
        } else {
            throw "Failed Acquiring Access Token!"
        }
    }
}
End {}
}

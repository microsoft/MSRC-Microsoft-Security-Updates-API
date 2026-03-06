Function Set-MSRCMsalAccessToken {
[CmdletBinding(SupportsShouldProcess)]
Param(
    [Parameter()]
    [Alias('ClientId')]
    [string]$ID,
    
    [Parameter()]
    [string]$TenantId = 'microsoft.onmicrosoft.com',
    
    [Parameter()]
    [string]$RedirectUri = 'http://localhost:50000'
)
Begin {}
Process {
    if ($PSCmdlet.ShouldProcess('Set the MSRCApiKey using MSRCMsalAccessToken')) {
        # Check if MSAL.PS module is available
        if (-not (Get-Module -ListAvailable -Name MSAL.PS)) {
            throw "MSAL.PS module is required. Please install it using: Install-Module -Name MSAL.PS"
        }
        
        Import-Module MSAL.PS -ErrorAction Stop
        
        # Clear any existing cached token
        $script:MSRCMsalAccessToken = $null
        
        try {
            # Use MSAL.PS to acquire token interactively
            $msalParams = @{
                ClientId    = $ID
                TenantId    = $TenantId
                Scopes      = @("$ID/.default")
                RedirectUri = $RedirectUri
                Interactive = $true
            }
            
            $tokenResult = Get-MsalToken @msalParams
            
            if ($tokenResult -and $tokenResult.AccessToken) {
                # Store only the access token string, not the full object
                $script:MSRCMsalAccessToken = @{
                    AccessToken = $tokenResult.AccessToken
                }
                Write-Verbose -Message "Successfully set your Access Token required by cmdlets of this module. Calls to the MSRC APIs will now use your access token."
            } else {
                throw "Failed Acquiring Access Token!"
            }
        }
        catch {
            throw
        }
    }
}
End {}
}

#Requires -Version 3.0

Function Set-MSRCApiKey {
[CmdletBinding()]
Param(
	[Parameter(Mandatory)]
	$ApiKey
)
Begin {}
Process {

	$global:MSRCApiKey = $ApiKey

    # we also set other shared variables
    $global:msrcApiUrl     = 'https://api.msrc.microsoft.com'
    $global:msrcApiVersion = 'api-version=2016-08-01'
}
End {}
}
Function Get-CVRFID {
[CmdletBinding()]
Param(
    [Parameter()]
    [Alias('CVRFID')]
    [string]$ID
)
Begin {
}
Process {
    if (-not ($global:MSRCApiKey)) {
	    Throw 'You need to use Set-MSRCApiKey first to set your API Key'

    } else {

        $url = '{0}/Updates?{1}' -f $global:msrcApiUrl,$global:msrcApiVersion

        try {
            if ($ID) {
                (Invoke-RestMethod -Uri $url -Headers @{'Accept' = 'application/json' ; 'Api-Key' = $global:MSRCApiKey } -ErrorAction Stop).Value | 
                Where { $_.ID -eq $ID }
    
            } else {
                ((Invoke-RestMethod -Uri $url -Headers @{'Accept' = 'application/json' ; 'Api-Key' = $global:MSRCApiKey } -ErrorAction Stop).Value).ID
            }
        } catch {
            Throw $_
        }
    }
}
End {}
}
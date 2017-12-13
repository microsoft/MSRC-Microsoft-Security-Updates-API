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
    $RestMethod = @{
        uri = '{0}/Updates?{1}' -f $global:msrcApiUrl,$global:msrcApiVersion
        Headers = @{
            'Accept' = 'application/json'
        }
        ErrorAction = 'Stop'
    }
    if ($global:msrcProxy){
        $RestMethod.Add('Proxy' , $global:msrcProxy)
    }
    if ($global:msrcProxyCredential){
        $RestMethod.Add('ProxyCredential',$global:msrcProxyCredential)
    }
    if ($global:MSRCApiKey) {
        
        $RestMethod.Headers.Add('Api-Key',$global:MSRCApiKey)
    
    } elseif ($global:MSRCAdalAccessToken) {
      
        $RestMethod.Headers.Add('Authorization',$($global:MSRCAdalAccessToken.CreateAuthorizationHeader()))

    } else {
    
        Throw 'You need to use Set-MSRCApiKey first to set your API Key'        
    }

    try {
    
        if ($ID) {

            (Invoke-RestMethod @RestMethod).Value | 
            Where-Object { $_.ID -eq $ID } | 
            Where-Object { $_ -ne '2017-May-B' }
    
        } else {
        
            ((Invoke-RestMethod @RestMethod).Value).ID | 
            Where-Object { $_ -ne '2017-May-B' }
        }

    } catch {
        
        Throw $_
    
    }
}
End {}
}
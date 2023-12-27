# we also set other shared variables
$global:msrcApiUrl     = 'https://api.msrc.microsoft.com/cvrf/v3.0'
Write-Verbose -Message "Successfully defined a msrcApiUrl global variable that points to $($global:msrcApiUrl)"

$global:msrcApiVersion = 'api-version=2023-11-01'
Write-Verbose -Message "Successfully defined a msrcApiVersion global variable that points to $($global:msrcApiVersion)"
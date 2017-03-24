Function Test-CVRFID {
[CmdletBinding()]
Param(
    [Parameter(Mandatory)]
    [ValidateNotNullorEmpty()]
    [Alias('CVRFID')]
    [string]$ID
)
Begin {}
Process {
    if (Get-CVRFID -ID $ID) {
        $true
    } else {
        $false
    }
}
End {}
}
Function Get-MsrcCvrfCVESummary {
<#
    .SYNOPSIS
        Get the CVE summary from vulnerabilities found in CVRF document

    .DESCRIPTION
       This function gathers the CVE Summary from vulnerabilities in a CVRF document.

    .PARAMETER Vulnerability
        A vulnerability object or objects from a CVRF document object

    .EXAMPLE
        Get-MsrcCvrfDocument -ID 2016-Nov | Get-MsrcCvrfCVESummary

        Get the CVE summary from a CVRF document using the pipeline.

    .EXAMPLE
        $cvrfDocument = Get-MsrcCvrfDocument -ID 2016-Nov
        Get-MsrcCvrfCVESummary -Vulnerability $cvrfDocument.Vulnerability

        Get the CVE summary from a CVRF document using a variable and parameters
#>
[CmdletBinding()]
Param (
    [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
    $Vulnerability,

    [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
    $ProductTree
)
Begin {

    $MaximumSeverityType = 3
    $ThreatsImpactType = 0

    Function Get-MaxSeverity {
    [CmdletBinding()]
    [OutputType('System.String')]
    Param($InputObject)
    Begin {}
    Process {
        if ('Critical' -in $InputObject) {
            'Critical'
        } elseif ('Important' -in $InputObject) {
            'Important'
        } elseif ('Moderate' -in $InputObject) {
            'Moderate'
        } elseif ('Low' -in $InputObject) {
            'Low'
        } else {
            'Unknown'
        }
    }
    End {}
    }
}
Process {

    $Vulnerability | ForEach-Object {

        $v = $_

        [PSCustomObject]@{
            CVE = $v.CVE
            Weakness = $(if ($v.CWE) { '{0} : {1}' -f "$($v.CWE.ID)","$($v.CWE.Value)"})
            Description = $(
                 ($v.Notes | Where-Object { $_.Title -eq 'Description' }).Value
            ) ;
            'Customer Action Required' = if ($customerActionNotes = $v.Notes | Where-Object { $_.Title -eq "Customer Action Required" }) {
                $customerActionNotes
            } else {
                'Yes'
            } ;
            'Maximum Severity Rating' = $(
                Get-MaxSeverity ($v.Threats | Where-Object {$_.Type -eq $MaximumSeverityType } ).Description.Value | Select-Object -Unique
            ) ;
            'Vulnerability Impact' = $(
                ($v.Threats | Where-Object {$_.Type -eq $ThreatsImpactType }).Description.Value | Select-Object -Unique
            ) ;
            'Affected Software' = $(
                $v.ProductStatuses.ProductID | ForEach-Object {
                    $id = $_
                    ($ProductTree.FullProductName | Where-Object { $_.ProductID -eq $id}).Value
                }
            ) ;
        }
    }
}
End {}
}
Function Get-MsrcCvrfAffectedSoftware {
<#
    .SYNOPSIS
        Get details of products affected by a CVRF document

    .DESCRIPTION
       CVRF documents next products into several places, including:
       -Vulnerabilities
       -Threats
       -Remediations
       -Product Tree
       This function gathers the details for each product identified in a CVRF document.

    .PARAMETER Vulnerability

    .PARAMETER ProductTree

    .EXAMPLE
        Get-MsrcCvrfDocument -ID 2016-Nov | Get-MsrcCvrfAffectedSoftware

        Get product details from a CVRF document using the pipeline.

    .EXAMPLE
        $cvrfDocument = Get-MsrcCvrfDocument -ID 2016-Nov
        Get-MsrcCvrfAffectedSoftware -Vulnerability $cvrfDocument.Vulnerability -ProductTree $cvrfDocument.ProductTree

        Get product details from a CVRF document using a variable and parameters
#>
[CmdletBinding()]
Param (
    [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
    $Vulnerability,

    [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
    $ProductTree
)
Begin {
    $MaximumSeverityType = Switch ("$($global:msrcApiUrl)") {
    'https://api.msrc.microsoft.com/cvrf/v3.0' {'Severity'}
    'https://api.msrc.microsoft.com/cvrf/v2.0' {3}
    }

    $ThreatsImpactType = Switch ("$($global:msrcApiUrl)") {
    'https://api.msrc.microsoft.com/cvrf/v3.0' {'Impact'}
    'https://api.msrc.microsoft.com/cvrf/v2.0' {0}
    }

    $RemediationsKBType = Switch ("$($global:msrcApiUrl)") {
    'https://api.msrc.microsoft.com/cvrf/v3.0' {'VendorFix'}
    'https://api.msrc.microsoft.com/cvrf/v2.0' {2}
    }

    $RemediationsKnownIssue = Switch ("$($global:msrcApiUrl)") {
    'https://api.msrc.microsoft.com/cvrf/v3.0' {'KnownIssue'}
    'https://api.msrc.microsoft.com/cvrf/v2.0' {5}
    }
}
Process {
    $Vulnerability | ForEach-Object {

        $v = $_

        $v.ProductStatuses.ProductID | ForEach-Object {
            $id = $_

            [PSCustomObject] @{
                FullProductName = $(
                    $ProductTree.FullProductName  |
                    Where-Object { $_.ProductID -eq $id} |
                    Select-Object -ExpandProperty Value
                ) ;
                KBArticle = $v.Remediations | where-Object {$_.ProductID -contains $id} | Where-Object {$_.Type -eq $RemediationsKBType} | ForEach-Object {
                                [PSCustomObject]@{
                                    ID = $_.Description.Value;
                                    URL= $_.URL;
                                    SubType = $_.SubType
                                }
                            };
               CVE = $v.CVE
               'Known Issue' = $v.Remediations | where-Object {$_.ProductID -contains $id} | Where-Object {$_.Type -eq $RemediationsKnownIssue} | ForEach-Object {
                                [PSCustomObject]@{
                                    ID = $_.Description.Value;
                                    URL= $_.URL;
                                }
               }
               Severity = $(
                    (
                        $v.Threats |
                        Where-Object {$_.Type -eq $MaximumSeverityType } |
                        Where-Object { $_.ProductID -contains $id }
                    ).Description.Value
               ) ;
               Impact = $(
                    (
                        $v.Threats |
                        Where-Object {$_.Type -eq $ThreatsImpactType } |
                        Where-Object { $_.ProductID -contains $id }
                    ).Description.Value
                );
               RestartRequired = $(
                    (
                        $v.Remediations |
                        Where-Object { $_.ProductID -contains $id }
                    ).RestartRequired.Value | ForEach-Object {
                        "$($_)"
                    }
               );
               FixedBuild = $(
                  (
                        $v.Remediations |
                        Where-Object { $_.ProductID -contains $id }
                    ).FixedBuild | ForEach-Object {
                        "$($_)"
                    }
               );
               Supercedence = $(
                    (
                        $v.Remediations |
                        Where-Object { $_.ProductID -contains $id }
                    ).Supercedence | ForEach-Object {
                        "$($_)"
                    }
               ) ;
               CvssScoreSet = $( [PSCustomObject]@{
                        base=    ($v.CVSSScoreSets | Where-Object { $_.ProductID -contains $id } | Select-Object -First 1).BaseScore;
                        temporal=($v.CVSSScoreSets | Where-Object { $_.ProductID -contains $id } | Select-Object -First 1).TemporalScore;
                        vector=  ($v.CVSSScoreSets | Where-Object { $_.ProductID -contains $id } | Select-Object -First 1).Vector;
                    }
               ) ;
            }
        }
    }
}
End {}
}

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
Begin {}
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
                KBArticle = $(
                    $v.Remediations | 
                    Where-Object { $_.ProductID -contains $id }| 
                    Select-Object -ExpandProperty Description  | 
                    Select-Object -ExpandProperty Value
                ) ;
                CVE = $v.CVE
                Severity = $(
                    $v.Threats | 
                    Where-Object {$_.Type -eq 3 } | 
                    Where-Object { $_.ProductID -contains $id } | 
                    Select-Object -ExpandProperty Description | 
                    Select-Object -ExpandProperty Value
                ) ;
                Impact = $(
                    $v.Threats | 
                    Where-Object {$_.Type -eq 0 } | 
                    Where-Object { $_.ProductID -contains $id } | 
                    Select-Object -ExpandProperty Description | 
                    Select-Object -ExpandProperty Value
                )
                RestartRequired = $(
                    (
                    $v.Remediations | 
                    Where-Object { $_.ProductID -contains $id }|
                    Select-Object -ExpandProperty RestartRequired |
                    Select-Object Value
                    ).Value | ForEach-Object {
                        if(-not($_)){
                            'Maybe'
                        } else {
                            "$($_)"
                        }
                    }
                ) ;

            }
        }
    }
}
End {}
}
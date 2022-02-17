Function Get-KBDownloadUrl {
<#
    .SYNOPSIS
        Takes the kb output from Get-MsrcCvrfAffectedSoftware and builds the html to insert into a document.

    .DESCRIPTION
        Takes the kb output from Get-MsrcCvrfAffectedSoftware and builds the html to insert into a document.

    .PARAMETER KBArticleObject
        The KB Article object that contains the id, url, and subtype.

    .EXAMPLE
        [PSCustomObject]@{ID="kb123456"; URL="microsoft.com"; SubType="Required"} | Get-KBDownloadUrl
#>
[CmdletBinding()]
[OutputType([System.String])]
Param (
    [Parameter(Mandatory,ValueFromPipeline)]
    [PSCustomObject]$KBArticleObject
)
Begin {
    $HTML_TO_RETURN = @()
}
Process {
    if (-not($KBArticleObject)){
        'None'
    } else {

        $KBArticleObject |
        ForEach-Object {
            $kb = $_
            #In older months, there won't be a subtype. Handle this so there are not empty ()'s
            if($kb.SubType){
                $HTML_TO_RETURN += $('<a href="{0}" >{1} ({2})' -f $kb.URL, $kb.ID, $kb.SubType)
            } else {
                $HTML_TO_RETURN += $('<a href="{0}" >{1}' -f $kb.URL, $kb.ID)
            }
        }
    }
}
End {
    $HTML_TO_RETURN -join '<br />'
}
}
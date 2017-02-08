#Define Rest API endpoints
$updateAllUrl  = "https://api.msrc.microsoft.com/Updates?api-Version=2016"
$updateByIDUrl = "https://api.msrc.microsoft.com/Updates('{0}')?api-version=2016-01-01"
$cvrfByIDUrl   = "https://api.msrc.microsoft.com/cvrf/{0}?api-version=2016-08-01"


#define functions for making api calls to the above URL's these will be called later in other functions that do more helpful things.
function GetAllUpdates
{
    if($args.Count -gt 0)
    {
        Write-Error "GetAllUpdates takes 0 arguments"
    }
    else
    {
        try
        {
            return (Invoke-RestMethod -Uri $updateAllUrl -ContentType application/json -Headers @{'Api-Key'="$apiKey"}).value
        } catch {
            Write-Error $("GetAllUpdates failed with status code {0}" -f $_.Exception.Response.StatusCode.value__ )
            Write-Error $_.Exception.Response.StatusDescription
        }
    }
}

function GetUpdateByID
{
    if($args.Count -ne 1)
    {
        Write-Error "GetUpdateByID takes 1 argument"
        Write-host "`tGetUpdateByID <id>"
        Write-host "`t"
        Write-host "`tID can contain one of the following:"
        Write-host "`t`t*Update ID        - ie. 2016-Aug"
        Write-host "`t`t*Vulnerability Id - ie. CVE-2016-0128"
        Write-host "`t`t*year             - ie. 2016"

    }
    else
    {
        try
        {
            return (Invoke-RestMethod -Uri $($updateByIDUrl -f $args[0]) -ContentType application/json -Headers @{'Api-Key'="$apiKey"}).value
        } catch {
            Write-Error $("GetUpdateByID failed with status code {0}" -f $_.Exception.Response.StatusCode.value__ )
            Write-Error $_.Exception.Response.StatusDescription
        }
    }
}

function GetCvrfByID
{
    if($args.Count -ne 1)
    {
        Write-Error "GetCvrfByID takes 1 argument"
        Write-host "`tGetCvrfByID <id>"
        Write-host "`t"
        Write-host "`tID can contain one of the following:"
        Write-host "`t`t*Update ID        - ie. 2016-Aug"
        Write-host "`t`t*Vulnerability Id - ie. CVE-2016-0128"
        Write-host "`t`t*year             - ie. 2016"

    }
    else
    {
        try
        {
            #by default this returns xml. set ContentType to get json output.
            $strResult = (Invoke-RestMethod -Uri $($cvrfByIDUrl -f $args[0]) -ContentType application/json -Headers @{'Api-Key'="$apiKey"})
            return $strResult
        } catch {
            Write-Error $("GetCvrfByID failed with status code {0}" -f $_.Exception.Response.StatusCode.value__)
            Write-Error $_.Exception.Response.StatusDescription
        }
    }
}


#A function that returns an array with all affected products, cve, cvss, and resources.
function GetAffectedProducts
{
    #make sure the number of args is correct.
    if($args.Count -ne 1)
    {
        Write-Error "GetAffectedProducts takes 1 argument"
        Write-host "`tGetAffectedProducts <id>"
        Write-host "`t"
        Write-host "`tID can contain one of the following:"
        Write-host "`t`t*Update ID        - ie. 2016-Aug"
        Write-host "`t`t*Vulnerability Id - ie. CVE-2016-0128"
        Write-host "`t`t*year             - ie. 2016"
        return

    }

    #get the json given by the endpoint
    $cvrfDoc = GetCvrfByID $args[0]

    #this is our array of objects which will contain the productID, productName, CVE, description, notes, and remediationUrls by the end of this function.
    $AffectedProducts = @()
    

    #First, we need to create mappings of all products to the cve

    #currently, there is duplicates in the Full product Name. this should change soon.
    foreach( $product in $($cvrfDoc.ProductTree.FullProductName | Sort-Object -Property ProductID | Get-Unique -AsString))
    {
        #for each vulnerability listed
        foreach( $vuln in $cvrfDoc.Vulnerability )
        {
            #for each productID group which was affected by the cve
            foreach( $productStatus in $vuln.ProductStatuses )
            {
                #for each product id in the productID group array
                foreach( $vulnProductId in $productStatus.ProductID)
                {
                    if( $vulnProductId -eq $product.ProductID )
                    {
                        $AffectedProducts += [pscustomobject]@{productID=$product.ProductID; productName=$product.Value; CVE=$vuln.CVE; description=$vuln.Title.Value; notes=$vuln.Notes.Value;
                        #also set the rest of the fields we will populate bloew to null, so if the api ever is missing a value, the field will still exist in the output
                        cvssBaseScore = $null;
                        cvssTemportalScore = $null;
                        cvssVector = $null;
                        remediationUrls = @();
                        Supercedence = $null
                        }
                    }
                }
            }
        }
    }


    #now we have product -> cve mapping, go back and add the cvss scores & remediations. 
    #the unique key for each cvss score is the (productID , CVE). 
    $AffectedProducts = foreach( $product in $AffectedProducts)
    {
        #find each vulnerability where the cve matches the cve thats affecting the product 
        foreach( $vuln in $cvrfDoc.Vulnerability | Where-Object {$_.CVE -eq $product.CVE})
        {
            #loop though all the cvss scores
            foreach( $cvssScore in $vuln.CVSSScoreSets )
            {
                #if the product ID matches the productID we are looking for, add the cvss information
                if($product.productID -eq $cvssScore.ProductID)
                {
                    $product.cvssBaseScore = $cvssScore.BaseScore
                    $product.cvssTemportalScore = $cvssScore.TemporalScore
                    $product.cvssVector = $cvssScore.Vector
                }
            }

            #since remediations and are inside of a vulnerability object, we can be a bit more efficient and look for the remediations here, and not have to look though all the vulnrabilities again later. 
            foreach( $remediation in $vuln.Remediations )
            {
                #loop though the productID's in the remediation object
                foreach( $remProductID in $remediation.ProductID )
                {
                    #if they match, we can add the remediation
                    if($product.productID -eq $remProductID)
                    {
                        #some remediations have empty urls. we want to skip those.

                        #if we already have a remediationUrls array, and the remediation URL is not empty, add it to the array 
                        if($remediation.URL -ne "")
                        {
                            $product.remediationUrls += $remediation.URL
                        }

                        #I havent seen any empty supercedence urls, but check anyway so we dont accedentilly remove the null with an empty string.
                        if($remediation.Supercedence -ne "")
                        {
                            $product.Supercedence += $remediation.Supercedence
                        }
                        
                    }
                }
            }
        }
        #place the updated product into the pipeline so we can set AffectedProducts to contain the new updated values.
        $product
    }

    return $AffectedProducts
}


#function to generate HTML reports.
function generateReport
{
    if($args.Count -ne 1)
    {
        Write-Error "generateReport takes 1 argument"
        Write-host "`tgenerateReport <id>"
        Write-host "`t"
        Write-host "`tID can contain one of the following:"
        Write-host "`t`t*Update ID        - ie. 2016-Aug"
        Write-host "`t`t*Vulnerability Id - ie. CVE-2016-0128"
        Write-host "`t`t*year             - ie. 2016"
        return
    }


    #get the product array and full document:
    $AffectedProducts = GetAffectedProducts $args[0]
    $cvrfDoc = GetCvrfByID $args[0]

    
    #define HTML templates with format strings:
    $HEADER    = "<h1>Generated Microsoft security bulletin summary for {0}</h1>" -f $($cvrfDoc.DocumentTracking.Identification.ID.Value)
    $SUBHEADER = "<h3>Generated - {0}</h3><br>" -f  $(Get-Date -Format g)
    
    $EXEC_SUMMARY_HEADER     = "<h2>Executive Summaries</h2><br><div id=`"execHeader`">"
    $EXEC_SUMMARY_TBL_HEADER = "<table border=`"1`"><tr><th>CVE</th><th>CVE Title</th> affected products</th><th>Affected Software</th></tr>"
    $EXEC_SUMMARY_TBL_BODY   = "<tr><td>{0}<br><a href=`"https://cve.mitre.org/cgi-bin/cvename.cgi?name={0}`">MITRE</a><br><a href=`"https://web.nvd.nist.gov/view/vuln/detail?vulnId={0}`">NVD</a></td><td>{1}</td><td>{2}</td></tr>"
    $EXEC_SUMMARY_TBL_FOOTER = "</table> </div><br><br>"

    $AFFECTED_SOFTWARE_HEADER        = "<h2>Affected Software</h2>"
    $AFFECTED_SOFTWARE_TBL_START    = "<table border=`"1`">"
    $AFFECTED_SOFTWARE_TBL_SW_HEADER = "<tr><th colspan = `"4`" style=`"background-color: #ededed`">{0}</th></tr>"
    $AFFECTED_SOFTWARE_TBL_BODY      = "<tr><td>{0}<br><a href=`"https://cve.mitre.org/cgi-bin/cvename.cgi?name={0}`">MITRE</a><br><a href=`"https://web.nvd.nist.gov/view/vuln/detail?vulnId={0}`">NVD</a></td><td>{1}</td><td><a href=`"{2}`">{2}</a></td><td>{3}</td></tr>" #in order: cve_id, cve_name, resources, cvss info
    $AFFECTED_SOFTWARE_FOOTER        = "</table><br>"
    #This is just the css used in traditional microsoft bulletins. feel free to change this to better match your needs 
    $CSS = "<link rel=`"stylesheet`" href=`"https://i-technet.sec.s-msft.com/Combined.css?resources=0:ImageSprite,0:TopicResponsive,0:TopicResponsive.MediaQueries,1:CodeSnippet,1:ProgrammingSelector,1:ExpandableCollapsibleArea,0:CommunityContent,1:TopicNotInScope,1:FeedViewerBasic,1:ImageSprite,2:Header.2,2:HeaderFooterSprite,2:Header.MediaQueries,2:Banner.MediaQueries,3:megabladeMenu.1,3:MegabladeMenu.MediaQueries,3:MegabladeMenuSpriteCluster,0:Breadcrumbs,0:Breadcrumbs.MediaQueries,0:ResponsiveToc,0:ResponsiveToc.MediaQueries,1:NavSidebar,0:LibraryMemberFilter,4:StandardRating,2:Footer.2,5:LinkList,2:Footer.MediaQueries,0:BaseResponsive,6:MsdnResponsive,0:Tables.MediaQueries,7:SkinnyRatingResponsive,7:SkinnyRatingV2;/Areas/Library/Content:0,/Areas/Epx/Content/Css:1,/Areas/Epx/Themes/TechNet/Content:2,/Areas/Epx/Themes/Shared/Content:3,/Areas/Global/Content:4,/Areas/Epx/Themes/Base/Content:5,/Areas/Library/Themes/Msdn/Content:6,/Areas/Library/Themes/TechNet/Content:7&amp;v=9192817066EC5D087D15C766A0430C95`">"
    #this is some simple css to limit the size of the "CVE Title" column in the executive summary section.
    $CSS2 = "<style>
                #execHeader td:first-child { width: 10% ;}
                #execHeader td:nth-child(3) { width: 35% ;}
            </style>"

    #build the document:

    #add the css in the html header, then add our headings to the doc
    $("<html><head>{0}</head><body>" -f ($CSS + $CSS2)) | Out-File 'report.html'

    #add a div to the entire document to bring it in from the sides a tad.
    "<div id=`"ducumentWrapper`" style=`"width: 90%; margin-left: auto; margin-right: auto;`">" >> "report.html"
    
    #add the header and subheader for the doc
    $HEADER    >> "report.html"
    $SUBHEADER >> "report.html"

    #begin the executive summary section, where our table is indexed by cve's
    $EXEC_SUMMARY_HEADER     >> "report.html"
    $EXEC_SUMMARY_TBL_HEADER >> "report.html"

    #loop though all the cve's for our summary
    foreach($vulnObject in $cvrfDoc.Vulnerability)
    {
        #create an array of affected products:
        $AffectedProductNames = $( $AffectedProducts | Where-Object {$_.CVE -eq $vulnObject.CVE}).productName | Sort-Object | Get-Unique
        $Resources            = $( $AffectedProducts | Where-Object {$_.CVE -eq $vulnObject.CVE}).remediationUrls | Sort-Object | Get-Unique
        
        #populate the table
        $($EXEC_SUMMARY_TBL_BODY -f (

                $vulnObject.CVE, #cve id for the first html form column
                $($vulnObject.Title.Value + "<br><hr>" + $vulnObject.Notes.Value), #cve name, second column
                $($AffectedProductNames -join "<br><br>")   #the list of products which are affected.
         )) >> "report.html"
    }

    #end the execurtive summary, and behin the affected software
    $EXEC_SUMMARY_TBL_FOOTER      >> "report.html"
    $AFFECTED_SOFTWARE_HEADER     >> "report.html"

    #add a div to keep the tables in a html object for styling
    "<div>" >> "report.html"

    #now for the more detailed affected products list
    foreach( $productName in $($AffectedProducts.productName | Sort-Object | Get-Unique ))
    {
        #make a new table for the affected peice of software
        $AFFECTED_SOFTWARE_TBL_START >> "report.html"

        #add the name of the software
        $($AFFECTED_SOFTWARE_TBL_SW_HEADER -f $productName) >> "report.html"

        #populate the table
        foreach( $vulnProductObj in $($AffectedProducts | Where-Object {$_.productName -eq $productName}))
        {
            $AFFECTED_SOFTWARE_TBL_BODY -f (
                $vulnProductObj.CVE , #cve id
                 $vulnProductObj.description , #cve description 
                 $($vulnProductObj.remediationUrls -join "<br><br>"), #remediation urls
                 $("cvss Base Score: " + $vulnProductObj.cvssBaseScore , "<br>cvss Temporal Score: " + $vulnProductObj.cvssTemportalScore , "<br>cvss vector: " + $vulnProductObj.cvssVector) #cvss info
            ) >> "report.html"

        }
        #end the table
        $AFFECTED_SOFTWARE_FOOTER >> "report.html"
    }
    #close up the affected software div and end the document.
    "</div>"               >> "report.html"
    "</div></body></html>" >> "report.html"
}






#ask user for api key:
$apiKey = Read-Host -Prompt "Please enter your api key"

#uncomment below to try out functions!
#generateReport "2016-Nov"
#GetAffectedProducts "2016-Nov"

#you can use all the great powershell cmdlets with GetAffectedProducts to quickly find objects of interest:
GetAffectedProducts "2016-Nov" | Where-Object {$_.productName -like "Internet Explorer 11 on Windows 8*"} | Where-Object {$_.cvssBaseScore -gt 4}
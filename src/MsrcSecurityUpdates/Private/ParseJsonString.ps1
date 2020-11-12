# ParseJsonString converts from string to PowerShell objects
# (workaround to overcome ConvertFrom-Json limitation on PowerShell 4.0 and earlier)

# Based on code by Florian Feldhaus at 'ConvertFrom-Json max length'
# https://stackoverflow.com/questions/16854057/convertfrom-json-max-length
# With the following changes:
# - ParseJsonString calls ParseItem, not ParseJsonObject (suggested by Dmitry Lobanov)
# - if-else test in ParseJsonObject is replaced by '$parsedItem = ParseItem $item'
# Note: this code is much slower than ConvertFrom-Json

Add-Type -Assembly System.Web.Extensions

# .NET JSON Serializer
$javaScriptSerializer = New-Object System.Web.Script.Serialization.JavaScriptSerializer
$javaScriptSerializer.MaxJsonLength = [System.Int32]::MaxValue
$javaScriptSerializer.RecursionLimit = 99

# Functions necessary to parse JSON output from .NET serializer to PowerShell Objects
function ParseItem($jsonItem) {
        if($jsonItem.PSObject.TypeNames -match "Array") {
                return ParseJsonArray($jsonItem)
        }
        elseif($jsonItem.PSObject.TypeNames -match "Dictionary") {
                return ParseJsonObject([HashTable]$jsonItem)
        }
        else {
                return $jsonItem
        }
}

function ParseJsonObject($jsonObj) {
        $result = New-Object -TypeName PSCustomObject
        foreach ($key in $jsonObj.Keys) {
                $item = $jsonObj[$key]
                $parsedItem = ParseItem $item
                $result | Add-Member -MemberType NoteProperty -Name $key -Value $parsedItem
        }
        return $result
}

function ParseJsonArray($jsonArray) {
        $result = @()
        $jsonArray | ForEach-Object {
                $result += , (ParseItem $_)
        }
        return $result
}

function ParseJsonString($json) {
        $config = $javaScriptSerializer.DeserializeObject($json)
        return ParseItem($config)
}

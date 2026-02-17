<#
.SYNOPSIS
    Converts Postman requests to Excel ODC (Office Data Connection) files for Power Query.

.DESCRIPTION
    This script reads a Postman collection or request export (JSON) and generates
    ODC files that can be imported into Excel Power Query. Supports:
    - GET and POST requests
    - Headers including Authorization
    - JSON request bodies
    - Variable substitution ({{variable}} syntax)
    - Optional pagination template for paginated APIs

.PARAMETER InputFile
    Path to the Postman collection JSON file or single request JSON file.

.PARAMETER OutputFolder
    Folder where ODC files will be created. Defaults to current directory.

.PARAMETER RequestName
    Name of specific request to convert from a collection. If not specified,
    converts the first request found or prompts for selection.

.PARAMETER Variables
    Hashtable of variable substitutions. E.g., @{baseUrl="https://api.example.com"; apiKey="abc123"}

.PARAMETER IncludePagination
    If specified, generates pagination boilerplate code for APIs that return paged results.

.PARAMETER PaginationTokenField
    The JSON field name for the next page token. Default: "next" (or "paging.next.after" for HubSpot-style)

.EXAMPLE
    .\Convert-PostmanToOdc.ps1 -InputFile "MyCollection.postman_collection.json" -RequestName "Get Users"

.EXAMPLE
    .\Convert-PostmanToOdc.ps1 -InputFile "request.json" -Variables @{apiKey="my-secret-key"} -IncludePagination
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputFile,

    [Parameter(Mandatory = $false)]
    [string]$OutputFolder = ".",

    [Parameter(Mandatory = $false)]
    [string]$RequestName,

    [Parameter(Mandatory = $false)]
    [hashtable]$Variables = @{},

    [Parameter(Mandatory = $false)]
    [switch]$IncludePagination,

    [Parameter(Mandatory = $false)]
    [string]$PaginationTokenField = "paging.next.after"
)

#region Helper Functions

function ConvertTo-HtmlEntities {
    <#
    .SYNOPSIS
        Converts special characters to HTML entities for ODC file embedding.
    #>
    param([string]$Text)
    
    $Text = $Text -replace '&', '&amp;'
    $Text = $Text -replace '<', '&lt;'
    $Text = $Text -replace '>', '&gt;'
    $Text = $Text -replace '"', '&quot;'
    return $Text
}

function Invoke-VariableSubstitution {
    <#
    .SYNOPSIS
        Replaces {{variableName}} placeholders with actual values.
    #>
    param(
        [string]$Text,
        [hashtable]$Variables
    )
    
    foreach ($key in $Variables.Keys) {
        $Text = $Text -replace "\{\{$key\}\}", $Variables[$key]
    }
    
    return $Text
}

function ConvertTo-PowerQueryRecord {
    <#
    .SYNOPSIS
        Converts a PowerShell hashtable to Power Query M record syntax.
    #>
    param([hashtable]$Hashtable)
    
    if ($Hashtable.Count -eq 0) {
        return "[]"
    }
    
    $entries = @()
    foreach ($key in $Hashtable.Keys) {
        $value = $Hashtable[$key]
        # Escape quotes in values
        $escapedValue = $value -replace '"', '""'
        $entries += "#`"$key`" = `"$escapedValue`""
    }
    
    return "[`n                $($entries -join ",`n                ")`n            ]"
}

function ConvertTo-PowerQueryValue {
    <#
    .SYNOPSIS
        Converts a JSON object/value to Power Query M literal syntax.
    #>
    param($Value, [int]$Indent = 0)
    
    $indentStr = " " * ($Indent * 4)
    
    if ($null -eq $Value) {
        return "null"
    }
    elseif ($Value -is [bool]) {
        return $Value.ToString().ToLower()
    }
    elseif ($Value -is [string]) {
        $escaped = $Value -replace '"', '""'
        return "`"$escaped`""
    }
    elseif ($Value -is [int] -or $Value -is [long] -or $Value -is [decimal] -or $Value -is [double]) {
        return $Value.ToString()
    }
    elseif ($Value -is [array]) {
        $items = @()
        foreach ($item in $Value) {
            $items += (ConvertTo-PowerQueryValue -Value $item -Indent ($Indent + 1))
        }
        return "{$($items -join ", ")}"
    }
    elseif ($Value -is [System.Collections.IDictionary] -or $Value.PSObject.Properties) {
        $entries = @()
        $props = if ($Value -is [System.Collections.IDictionary]) { $Value.Keys } else { $Value.PSObject.Properties.Name }
        foreach ($key in $props) {
            $propValue = if ($Value -is [System.Collections.IDictionary]) { $Value[$key] } else { $Value.$key }
            $entries += "$key = $(ConvertTo-PowerQueryValue -Value $propValue -Indent ($Indent + 1))"
        }
        return "[$($entries -join ", ")]"
    }
    else {
        return "`"$Value`""
    }
}

function Get-PostmanRequest {
    <#
    .SYNOPSIS
        Extracts a request from a Postman collection or request file.
    #>
    param(
        [object]$JsonContent,
        [string]$RequestName
    )
    
    # Check if this is a collection (has "item" property) or single request
    if ($JsonContent.item) {
        # It's a collection - find the request
        $requests = @()
        
        function Find-Requests {
            param($Items, [string]$ParentPath = "")
            
            foreach ($item in $Items) {
                $currentPath = if ($ParentPath) { "$ParentPath / $($item.name)" } else { $item.name }
                
                if ($item.request) {
                    # This is a request
                    [PSCustomObject]@{
                        Name = $item.name
                        Path = $currentPath
                        Request = $item.request
                        Response = $item.response
                    }
                }
                elseif ($item.item) {
                    # This is a folder - recurse
                    Find-Requests -Items $item.item -ParentPath $currentPath
                }
            }
        }
        
        $requests = @(Find-Requests -Items $JsonContent.item)
        
        if ($requests.Count -eq 0) {
            throw "No requests found in collection"
        }
        
        if ($RequestName) {
            $selected = $requests | Where-Object { $_.Name -eq $RequestName -or $_.Path -eq $RequestName }
            if (-not $selected) {
                throw "Request '$RequestName' not found. Available requests:`n$($requests.Name -join "`n")"
            }
            return $selected | Select-Object -First 1
        }
        
        # Return first request if only one, otherwise list them
        if ($requests.Count -eq 1) {
            return $requests[0]
        }
        
        Write-Host "Multiple requests found. Please specify -RequestName:" -ForegroundColor Yellow
        $requests | ForEach-Object { Write-Host "  - $($_.Path)" }
        throw "Multiple requests found. Specify -RequestName parameter."
    }
    elseif ($JsonContent.request) {
        # Single request format
        $reqName = if ($JsonContent.name) { $JsonContent.name } else { "Request" }
        return [PSCustomObject]@{
            Name = $reqName
            Path = $reqName
            Request = $JsonContent.request
            Response = $JsonContent.response
        }
    }
    elseif ($JsonContent.url -or $JsonContent.method) {
        # Direct request object
        return [PSCustomObject]@{
            Name = "Request"
            Path = "Request"
            Request = $JsonContent
            Response = $null
        }
    }
    else {
        throw "Unrecognized Postman JSON format"
    }
}

function Build-RequestUrl {
    <#
    .SYNOPSIS
        Builds URL from Postman URL object (can be string or object with host/path/query).
    #>
    param($UrlObject)
    
    if ($UrlObject -is [string]) {
        return $UrlObject
    }
    
    $url = ""
    
    # Protocol
    $protocol = if ($UrlObject.protocol) { $UrlObject.protocol } else { "https" }
    $url = "$protocol`://"
    
    # Host
    if ($UrlObject.host -is [array]) {
        $url += $UrlObject.host -join "."
    }
    else {
        $url += $UrlObject.host
    }
    
    # Path
    if ($UrlObject.path) {
        if ($UrlObject.path -is [array]) {
            $url += "/" + ($UrlObject.path -join "/")
        }
        else {
            $url += "/" + $UrlObject.path
        }
    }
    
    # Query parameters
    if ($UrlObject.query -and $UrlObject.query.Count -gt 0) {
        $queryParams = @()
        foreach ($param in $UrlObject.query) {
            if (-not $param.disabled) {
                $queryParams += "$($param.key)=$($param.value)"
            }
        }
        if ($queryParams.Count -gt 0) {
            $url += "?" + ($queryParams -join "&")
        }
    }
    
    return $url
}

function Build-HeadersRecord {
    <#
    .SYNOPSIS
        Builds Power Query headers record from Postman headers array.
    #>
    param($Headers)
    
    $headerDict = @{}
    
    if ($Headers) {
        foreach ($header in $Headers) {
            if (-not $header.disabled) {
                $headerDict[$header.key] = $header.value
            }
        }
    }
    
    # Ensure Content-Type is set for requests with body
    if (-not $headerDict.ContainsKey("Content-Type")) {
        $headerDict["Content-Type"] = "application/json"
    }
    
    return $headerDict
}

function Convert-PostmanBodyToPowerQuery {
    <#
    .SYNOPSIS
        Converts Postman request body to Power Query M code.
    #>
    param($Body)
    
    if (-not $Body) {
        return $null
    }
    
    switch ($Body.mode) {
        "raw" {
            # Try to parse as JSON for better M code generation
            try {
                $jsonBody = $Body.raw | ConvertFrom-Json
                $mValue = ConvertTo-PowerQueryValue -Value $jsonBody
                return "Text.ToBinary(Json.FromValue($mValue))"
            }
            catch {
                # Not JSON, use raw text
                $escaped = $Body.raw -replace '"', '""'
                return "Text.ToBinary(`"$escaped`")"
            }
        }
        "formdata" {
            # Form data - build as record
            $formEntries = @()
            foreach ($item in $Body.formdata) {
                if (-not $item.disabled) {
                    $escaped = $item.value -replace '"', '""'
                    $formEntries += "$($item.key) = `"$escaped`""
                }
            }
            return "Text.ToBinary(Uri.BuildQueryString([$($formEntries -join ", ")]))"
        }
        "urlencoded" {
            # URL encoded - similar to formdata
            $formEntries = @()
            foreach ($item in $Body.urlencoded) {
                if (-not $item.disabled) {
                    $escaped = $item.value -replace '"', '""'
                    $formEntries += "$($item.key) = `"$escaped`""
                }
            }
            return "Text.ToBinary(Uri.BuildQueryString([$($formEntries -join ", ")]))"
        }
        default {
            return $null
        }
    }
}

function New-PowerQueryCode {
    <#
    .SYNOPSIS
        Generates Power Query M code for a web request.
    #>
    param(
        [string]$Url,
        [string]$Method,
        [hashtable]$Headers,
        [string]$BodyCode,
        [bool]$IncludePagination,
        [string]$PaginationTokenField,
        [string]$QueryName
    )
    
    $headersRecord = ConvertTo-PowerQueryRecord -Hashtable $Headers
    
    # Simple request without pagination
    if (-not $IncludePagination) {
        $webContentsOptions = @()
        $webContentsOptions += "Headers = $headersRecord"
        
        if ($BodyCode) {
            $webContentsOptions += "Content = $BodyCode"
        }
        
        $optionsCode = "[`n            $($webContentsOptions -join ",`n            ")`n        ]"
        
        $mCode = @"
let
    // API Configuration
    apiUrl = "$Url",
    
    // Make the API request
    response = Json.Document(Web.Contents(apiUrl, $optionsCode)),
    
    // Convert to table (adjust based on your response structure)
    result = response
in
    result
"@
    }
    else {
        # With pagination support
        $webContentsOptions = @()
        $webContentsOptions += "Headers = headers"
        
        if ($BodyCode) {
            # For pagination, we need to handle the body differently to include afterToken
            $webContentsOptions += "Content = requestBody"
        }
        
        $optionsCode = "[`n                $($webContentsOptions -join ",`n                ")`n            ]"
        
        # Parse pagination field path (e.g., "paging.next.after" -> response[paging][next][after])
        $paginationParts = $PaginationTokenField -split '\.'
        $paginationAccessor = ($paginationParts | ForEach-Object { "[$_]" }) -join ""
        
        $mCode = @"
let
    // API Configuration
    apiUrl = "$Url",
    headers = $headersRecord,
    
    // Function to make the API call with pagination support
    GetPage = (afterToken as nullable text) =>
        let
            // Build request body (modify as needed for your API's pagination parameter)
            requestBody = Text.ToBinary(Json.FromValue([
                // Add your request parameters here
                // For pagination, include: after = if afterToken <> null then afterToken else null
            ])),
            
            response = Json.Document(Web.Contents(apiUrl, $optionsCode)),
            results = response[results],
            nextToken = try response$paginationAccessor otherwise null,
            output = [Results = results, NextPage = nextToken]
        in
            output,
    
    // Initialize first page
    initialResponse = GetPage(null),
    allResults = initialResponse[Results],
    nextPageToken = initialResponse[NextPage],
    
    // Function to iterate through all pages
    GetAllPages = (currentResults, pageToken) =>
        let
            newResponse = if pageToken <> null then GetPage(pageToken) else [Results = {}, NextPage = null],
            newResults = newResponse[Results],
            combinedResults = List.Combine({currentResults, newResults}),
            newNextPageToken = newResponse[NextPage]
        in
            if newNextPageToken <> null then @GetAllPages(combinedResults, newNextPageToken) else combinedResults,
    
    // Get all results
    allData = GetAllPages(allResults, nextPageToken),
    
    // Convert to table
    dataTable = Table.FromList(allData, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    
    // Expand the results (adjust column names based on your API response)
    result = if Table.HasColumns(dataTable, "Column1") then
        Table.ExpandRecordColumn(dataTable, "Column1", Record.FieldNames(dataTable{0}[Column1]))
    else
        dataTable
in
    result
"@
    }
    
    return $mCode
}

function New-OdcFile {
    <#
    .SYNOPSIS
        Creates an ODC file with the specified Power Query code.
    #>
    param(
        [string]$QueryName,
        [string]$Description,
        [string]$PowerQueryCode
    )
    
    # IMPORTANT: Order matters for encoding!
    # 1. First HTML-encode special chars in the M code (& < > ")
    # 2. Then replace newlines with &#13;&#10; (so the & in these doesn't get escaped)
    $mashupCode = $PowerQueryCode -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;'
    $mashupCode = $mashupCode -replace "`r`n", "&#13;&#10;" -replace "`n", "&#13;&#10;"
    
    # Build the mashup XML (pre-encoded for HTML embedding)
    $mashupXml = @"
&lt;Mashup xmlns:xsd=&quot;http://www.w3.org/2001/XMLSchema&quot; xmlns:xsi=&quot;http://www.w3.org/2001/XMLSchema-instance&quot; xmlns=&quot;http://schemas.microsoft.com/DataMashup&quot;&gt;&lt;Client&gt;EXCEL&lt;/Client&gt;&lt;Version&gt;2.129.252.0&lt;/Version&gt;&lt;MinVersion&gt;2.21.0.0&lt;/MinVersion&gt;&lt;Culture&gt;en-GB&lt;/Culture&gt;&lt;SafeCombine&gt;false&lt;/SafeCombine&gt;&lt;Items&gt;&lt;Query Name=&quot;$QueryName&quot;&gt;&lt;Formula&gt;&lt;![CDATA[$mashupCode]]&gt;&lt;/Formula&gt;&lt;IsParameterQuery xsi:nil=&quot;true&quot; /&gt;&lt;IsDirectQuery xsi:nil=&quot;true&quot; /&gt;&lt;/Query&gt;&lt;/Items&gt;&lt;/Mashup&gt;
"@

    $odcContent = @"
<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/x-ms-odc; charset=utf-8">
<meta name=ProgId content=ODC.Database>
<meta name=SourceType content=OLEDB>
<title>Query - $QueryName</title>
<xml id=docprops><o:DocumentProperties
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns="http://www.w3.org/TR/REC-html40">
  <o:Description>$Description</o:Description>
  <o:Name>Query - $QueryName</o:Name>
 </o:DocumentProperties>
</xml><xml id=msodc><odc:OfficeDataConnection
  xmlns:odc="urn:schemas-microsoft-com:office:odc"
  xmlns="http://www.w3.org/TR/REC-html40">
  <odc:PowerQueryConnection odc:Type="OLEDB">
   <odc:ConnectionString>Provider=Microsoft.Mashup.OleDb.1;Data Source=`$Workbook`$;Location=$QueryName;Extended Properties=&quot;&quot;</odc:ConnectionString>
   <odc:CommandType>SQL</odc:CommandType>
   <odc:CommandText>SELECT * FROM [$QueryName]</odc:CommandText>
  </odc:PowerQueryConnection>
  <odc:PowerQueryMashupData>$mashupXml</odc:PowerQueryMashupData>
 </odc:OfficeDataConnection>
</xml>
<style>
<!--
    .ODCDataSource
    {
    behavior: url(dataconn.htc);
    }
-->
</style>
 
</head>

<body onload='init()' scroll=no leftmargin=0 topmargin=0 rightmargin=0 style='border: 0px'>
<table style='border: solid 1px threedface; height: 100%; width: 100%' cellpadding=0 cellspacing=0 width='100%'> 
  <tr> 
    <td id=tdName style='font-family:arial; font-size:medium; padding: 3px; background-color: threedface'> 
      &nbsp; 
    </td> 
     <td id=tdTableDropdown style='padding: 3px; background-color: threedface; vertical-align: top; padding-bottom: 3px'>

      &nbsp; 
    </td> 
  </tr> 
  <tr> 
    <td id=tdDesc colspan='2' style='border-bottom: 1px threedshadow solid; font-family: Arial; font-size: 1pt; padding: 2px; background-color: threedface'>

      &nbsp; 
    </td> 
  </tr> 
  <tr> 
    <td colspan='2' style='height: 100%; padding-bottom: 4px; border-top: 1px threedhighlight solid;'> 
      <div id='pt' style='height: 100%' class='ODCDataSource'></div> 
    </td> 
  </tr> 
</table> 

  
<script language='javascript'> 

function init() { 
  var sName, sDescription; 
  var i, j; 
  
  try { 
    sName = unescape(location.href) 
  
    i = sName.lastIndexOf(".") 
    if (i>=0) { sName = sName.substring(1, i); } 
  
    i = sName.lastIndexOf("/") 
    if (i>=0) { sName = sName.substring(i+1, sName.length); } 

    document.title = sName; 
    document.getElementById("tdName").innerText = sName; 

    sDescription = document.getElementById("docprops").innerHTML; 
  
    i = sDescription.indexOf("escription>") 
    if (i>=0) { j = sDescription.indexOf("escription>", i + 11); } 

    if (i>=0 && j >= 0) { 
      j = sDescription.lastIndexOf("</", j); 

      if (j>=0) { 
          sDescription = sDescription.substring(i+11, j); 
        if (sDescription != "") { 
            document.getElementById("tdDesc").style.fontSize="x-small"; 
          document.getElementById("tdDesc").innerHTML = sDescription; 
          } 
        } 
      } 
    } 
  catch(e) { 

    } 
  } 
</script> 

</body> 
 
</html>
"@

    return $odcContent
}

#endregion

#region Main Script

try {
    # Strip quotes from paths (in case user entered them at prompt)
    $InputFile = $InputFile.Trim('"', "'", ' ')
    $OutputFolder = $OutputFolder.Trim('"', "'", ' ')
    
    # Validate input file
    if (-not (Test-Path $InputFile)) {
        throw "Input file not found: $InputFile"
    }
    
    # Create output folder if needed
    if (-not (Test-Path $OutputFolder)) {
        New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
    }
    
    # Read and parse the Postman JSON
    Write-Host "Reading Postman file: $InputFile" -ForegroundColor Cyan
    $jsonContent = Get-Content $InputFile -Raw | ConvertFrom-Json
    
    # Get the request
    $requestInfo = Get-PostmanRequest -JsonContent $jsonContent -RequestName $RequestName
    Write-Host "Processing request: $($requestInfo.Name)" -ForegroundColor Cyan
    
    $request = $requestInfo.Request
    
    # Build URL
    $url = Build-RequestUrl -UrlObject $request.url
    $url = Invoke-VariableSubstitution -Text $url -Variables $Variables
    Write-Host "  URL: $url" -ForegroundColor Gray
    
    # Get HTTP method
    $method = if ($request.method) { $request.method.ToUpper() } else { "GET" }
    Write-Host "  Method: $method" -ForegroundColor Gray
    
    # Build headers
    $headers = Build-HeadersRecord -Headers $request.header
    foreach ($key in @($headers.Keys)) {
        $headers[$key] = Invoke-VariableSubstitution -Text $headers[$key] -Variables $Variables
    }
    Write-Host "  Headers: $($headers.Count) configured" -ForegroundColor Gray
    
    # Build body code
    $bodyCode = $null
    if ($method -in @("POST", "PUT", "PATCH") -and $request.body) {
        $bodyRaw = $request.body.raw
        if ($bodyRaw) {
            $bodyRaw = Invoke-VariableSubstitution -Text $bodyRaw -Variables $Variables
            $request.body.raw = $bodyRaw
        }
        $bodyCode = Convert-PostmanBodyToPowerQuery -Body $request.body
        Write-Host "  Body: Included" -ForegroundColor Gray
    }
    
    # Generate Power Query code
    Write-Host "`nGenerating Power Query code..." -ForegroundColor Cyan
    $pqCode = New-PowerQueryCode `
        -Url $url `
        -Method $method `
        -Headers $headers `
        -BodyCode $bodyCode `
        -IncludePagination $IncludePagination.IsPresent `
        -PaginationTokenField $PaginationTokenField `
        -QueryName $requestInfo.Name
    
    # Create ODC file
    $safeName = $requestInfo.Name -replace '[\\/:*?"<>|]', '_'
    $odcContent = New-OdcFile `
        -QueryName $safeName `
        -Description "Converted from Postman request: $($requestInfo.Name)" `
        -PowerQueryCode $pqCode
    
    # Write ODC file
    $outputPath = Join-Path $OutputFolder "Query - $safeName.odc"
    Set-Content -Path $outputPath -Value $odcContent -Encoding UTF8
    
    Write-Host "`nODC file created successfully!" -ForegroundColor Green
    Write-Host "Output: $outputPath" -ForegroundColor Gray
    
    # Also output the M code to console for reference
    Write-Host "`n--- Generated Power Query M Code ---" -ForegroundColor Yellow
    Write-Host $pqCode
    Write-Host "--- End of M Code ---`n" -ForegroundColor Yellow
    
    return $outputPath
}
catch {
    Write-Error "Error: $_"
    throw
}

#endregion

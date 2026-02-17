# Postman to ODC Converter

Convert Postman API requests to Excel ODC (Office Data Connection) files for use with Power Query.

## Overview

This tool allows you to:
1. Export a request from Postman as JSON
2. Run the converter script
3. Get an `.odc` file ready to import into Excel Power Query

## Quick Start

### Basic Usage

```powershell
# Convert a simple request
.\Convert-PostmanToOdc.ps1 -InputFile "my-request.json"

# Convert a specific request from a collection
.\Convert-PostmanToOdc.ps1 -InputFile "MyCollection.postman_collection.json" -RequestName "Get Users"

# With variable substitution
.\Convert-PostmanToOdc.ps1 -InputFile "my-request.json" -Variables @{
    apiKey = "your-api-key"
    baseUrl = "https://api.example.com"
}

# With pagination support
.\Convert-PostmanToOdc.ps1 -InputFile "my-request.json" -IncludePagination -PaginationTokenField "paging.next.after"
```

### How to Export from Postman

1. **Single Request**: Right-click on a request → Export → Save as JSON
2. **Collection**: Click the "..." on a collection → Export → Collection v2.1 (recommended)

## Parameters

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-InputFile` | Yes | Path to Postman JSON file |
| `-OutputFolder` | No | Where to save ODC files (default: current dir) |
| `-RequestName` | No | Name of request to convert from a collection |
| `-Variables` | No | Hashtable of `{{variable}}` substitutions |
| `-IncludePagination` | No | Generate pagination boilerplate code |
| `-PaginationTokenField` | No | JSON path for next page token (default: `paging.next.after`) |

## Variable Substitution

Postman variables like `{{apiKey}}` can be replaced with actual values:

```powershell
.\Convert-PostmanToOdc.ps1 -InputFile "request.json" -Variables @{
    apiKey = "pat-na1-xxxxx"
    startDate = "1704067200000"
    endDate = "1704672000000"
}
```

**Note**: For sensitive values like API keys, consider editing the generated ODC file to reference Excel named ranges instead of hardcoding.

## Examples

### Example 1: Simple GET Request

Postman request:
```json
{
    "name": "Get Users",
    "request": {
        "method": "GET",
        "url": "https://api.example.com/users",
        "header": [
            { "key": "Authorization", "value": "Bearer {{apiKey}}" }
        ]
    }
}
```

Convert:
```powershell
.\Convert-PostmanToOdc.ps1 -InputFile "get-users.json" -Variables @{apiKey="my-secret"}
```

### Example 2: POST with Pagination (HubSpot-style)

```powershell
.\Convert-PostmanToOdc.ps1 `
    -InputFile "examples\sample-hubspot-request.json" `
    -RequestName "Search Contacts" `
    -Variables @{
        hubspot_api_key = "pat-na1-xxxxx"
        startDate = "1704067200000"
        endDate = "1704672000000"
    } `
    -IncludePagination
```

## Using the Generated ODC File

1. **In Excel**: Data → Get Data → From File → From Text/CSV → Select the `.odc` file
2. **Or**: Double-click the `.odc` file to open in Excel
3. **Or**: Copy the query into Power Query Editor directly

## Supported Postman Features

| Feature | Supported |
|---------|-----------|
| GET requests | ✅ |
| POST/PUT/PATCH requests | ✅ |
| Headers | ✅ |
| JSON body | ✅ |
| Form data | ✅ |
| URL parameters | ✅ |
| Variables `{{var}}` | ✅ |
| Bearer auth | ✅ |
| Basic auth | ⚠️ (via headers) |
| Collections | ✅ |
| Folders in collections | ✅ |
| Pre-request scripts | ❌ |
| Tests/Post scripts | ❌ |

## Troubleshooting

### "Multiple requests found"
Specify which request to convert with `-RequestName`:
```powershell
.\Convert-PostmanToOdc.ps1 -InputFile "Collection.json" -RequestName "Search Contacts"
```

### Variables not replaced
Ensure variable names match exactly (case-sensitive):
```powershell
# If Postman uses {{API_KEY}}, use:
-Variables @{API_KEY = "value"}  # Not @{apiKey = "value"}
```

### ODC file won't import
- Ensure the file has `.odc` extension
- Try importing via Data → Get Data → From File → From Text/CSV

## Advanced: Modifying Generated Queries

The generated M code can be customized:

1. Open the ODC file in a text editor
2. Find the `<odc:PowerQueryMashupData>` section
3. The M code is HTML-encoded within `<![CDATA[...]]>`
4. Or import into Excel and edit in Power Query Editor

### Common Modifications

**Reference Excel parameters instead of hardcoded values:**
```m
// Change from:
apiKey = "hardcoded-value",

// To:
apiKey = Excel.CurrentWorkbook(){[Name="Settings"]}[Content]{0}[ApiKey],
```

**Adjust pagination for different APIs:**
```m
// For cursor-based pagination:
nextToken = try response[meta][next_cursor] otherwise null

// For offset pagination:
nextOffset = try response[offset] + response[limit] otherwise null
```

## License

MIT License - Feel free to modify and distribute.

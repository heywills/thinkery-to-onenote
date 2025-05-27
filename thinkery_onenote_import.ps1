<#
.SYNOPSIS
    Import Thinkery JSON export into a new OneNote notebook via Microsoft Graph.

.DESCRIPTION
    1. Creates a new notebook (default name "Thinkery Import", override with -NotebookName).
    2. Builds section groups and sections according to the mapping discussed with ChatGPT.
    3. Parses the Thinkery JSON export, creating:
       * One page per "large" note (≥ 300 characters).
       * Aggregated "Quick Notes – {Topic}" pages, where every tiny note (< 300 characters) becomes its own <h3> heading + body.
    4. Requires a short‑lived delegated access token with Notes.ReadWrite scope.

.PARAMETER AccessToken
    OAuth 2.0 bearer token copied from Graph Explorer (Notes.ReadWrite).

.PARAMETER JsonPath
    Path to the thinkery‑tiriansdoor.json file.

.PARAMETER NotebookName
    Display name for the new notebook.  Default: "Thinkery Import"

.PARAMETER ImportMapPath
    Path to the JSON file defining OneNote structure and tag mappings.
    Default: "./sample-import-maps/heywills-import-map.json"

.PARAMETER TinyNoteThreshold
    Character count threshold below which notes are considered "tiny" and will be aggregated.
    Default: 140

.PARAMETER DryRun
    If specified, the script will not make any changes, only report what it would do.

.INSTRUCTIONS
    1. Open Graph Explorer: https://developer.microsoft.com/en-us/graph/graph-explorer
    2. Sign in with your Microsoft account.
    3. Set the permissions to Notes.ReadWrite.
       a.Choose API Explorer from the left menu.
       b. Expand me->onenote->notebooks->post
       c. Click "Modify permissions"
       d. Click Consent next to "Notes.ReadWrite".
    4. Click the "Access token" section and copy the token.
    5. Run this script with the copied token and your Thinkery JSON export file.
    Example:
        .\thinkery_onenote_import.ps1 -AccessToken "eyJ0eXAiOiJKV
#>

param(
    [Parameter(Mandatory = $true)][string]$AccessToken,
    [string]$JsonPath = ".\\import-files\\thinkery-tiriansdoor.json",
    [string]$NotebookName = "Thinkery Tiriansdoor Import",
    [string]$ImportMapPath = ".\\sample-import-maps\\heywills-import-map.json",
    [int]$TinyNoteThreshold = 140,
    [switch]$DryRun = $false
)

$ErrorActionPreference = "Break"
$graphApi = "https://graph.microsoft.com/v1.0"

Function Invoke-GraphPost($Uri, $BodyObj) {
    $json = $BodyObj | ConvertTo-Json -Depth 6
    try {
        # For debugging
        Write-Debug "Sending request to $Uri with body: $json"
        
        if ($DryRun) {
            Write-Host "[DRY RUN] Would send request to $Uri" -ForegroundColor Yellow
            return [PSCustomObject]@{ id = "dry-run-id-$(Get-Random)" }
        }
        
        $response = Invoke-RestMethod -Method Post -Uri $Uri `
            -Headers @{ Authorization = "Bearer $AccessToken"; "Content-Type" = "application/json" } `
            -Body $json -ErrorVariable responseError
        return $response
    } catch {
        Write-Error "Graph API Error: $_"
        Write-Error "Request body: $json"
        throw $_
    }
}

Function Sanitize-Name {
    param([string]$Name)
    # Replace problematic characters
    $sanitized = $Name -replace "&", "and" `
                      -replace "\+", "plus" `
                      -replace "#", "num" `
                      -replace "%", "percent" `
                      -replace "/", "-" 
    return $sanitized
}

Function Create-Notebook {
    param([string]$Name)
    $sanitizedName = Sanitize-Name -Name $Name
    Write-Host "Creating notebook '$Name' (sanitized as '$sanitizedName')..."
    $nb = Invoke-GraphPost "$graphApi/me/onenote/notebooks" @{ displayName = $sanitizedName }
    return $nb.id
}

Function Create-SectionGroup {
    param([string]$NotebookId, [string]$Name)
    $sanitizedName = Sanitize-Name -Name $Name
    Write-Host "Creating section group '$Name' (sanitized as '$sanitizedName')..."
    $sg = Invoke-GraphPost "$graphApi/me/onenote/notebooks/$NotebookId/sectionGroups" @{ displayName = $sanitizedName }
    return $sg.id
}

Function Create-Section {
    param([string]$SectionGroupId, [string]$Name)
    $sanitizedName = Sanitize-Name -Name $Name
    Write-Host "Creating section '$Name' (sanitized as '$sanitizedName')..."
    $sec = Invoke-GraphPost "$graphApi/me/onenote/sectionGroups/$SectionGroupId/sections" @{ displayName = $sanitizedName }
    return $sec.id
}

Function Post-Page {
    param([string]$SectionId, [string]$Html)
    try {
        if ($DryRun) {
            Write-Host "[DRY RUN] Would post page to section $SectionId" -ForegroundColor Yellow
            return
        }
        
        # Send the HTML content directly in the request body
        $ret = Invoke-RestMethod -Method Post -Uri "$graphApi/me/onenote/sections/$SectionId/pages" `
            -Headers @{ 
                "Authorization" = "Bearer $AccessToken"
                "Content-Type" = "text/html; charset=utf-8"
            } `
            -Body $Html
    }
    catch {
        Write-Error "Error posting page: $_"
        Break
    }
}

# Function to validate the import map structure
Function Validate-ImportMap($importMap) {
    if ($null -eq $importMap) {
        throw "Import map is null or could not be loaded."
    }

    if ($importMap -isnot [array]) {
        throw "Import map must be an array."
    }

    foreach ($group in $importMap) {
        # Validate group has required properties
        if (-not $group.ContainsKey('OneNoteSectionGroupName')) {
            throw "Each group in the import map must have a OneNoteSectionGroupName property."
        }

        if ($group.OneNoteSectionGroupName -isnot [string]) {
            throw "OneNoteSectionGroupName must be a string."
        }

        if (-not $group.ContainsKey('OneNoteSections')) {
            throw "Each group in the import map must have a OneNoteSections property."
        }

        if ($group.OneNoteSections -isnot [array]) {
            throw "OneNoteSections must be an array."
        }

        # Validate each section has required properties
        foreach ($section in $group.OneNoteSections) {
            if (-not $section.ContainsKey('OneNoteSectionName')) {
                throw "Each section must have a OneNoteSectionName property."
            }

            if ($section.OneNoteSectionName -isnot [string]) {
                throw "OneNoteSectionName must be a string."
            }

            if (-not $section.ContainsKey('ThinkeryTags')) {
                throw "Each section must have a ThinkeryTags property."
            }

            if ($section.ThinkeryTags -isnot [array]) {
                throw "ThinkeryTags must be an array."
            }

            # Ensure all tags are strings
            foreach ($tag in $section.ThinkeryTags) {
                if ($tag -isnot [string]) {
                    throw "Each tag in ThinkeryTags must be a string."
                }
            }
        }
    }

    return $true
}

# Constants for default uncategorized content
$DEFAULT_GROUP_NAME = "Uncategorized"
$DEFAULT_SECTION_NAME = "Uncategorized imported items"

# 1. Notebook
if ($DryRun) {
    Write-Host "[DRY RUN] Would create notebook '$NotebookName'" -ForegroundColor Yellow
    $notebookId = "dry-run-notebook-id"
} else {
    $notebookId = Create-Notebook -Name $NotebookName
}
Write-Host "Notebook created with id $notebookId"

# 2. Load and validate import map
try {
    Write-Host "Loading import map from $ImportMapPath..."
    
    if (-not (Test-Path $ImportMapPath)) {
        throw "Import map file not found: $ImportMapPath"
    }
    
    $importMap = Get-Content $ImportMapPath -Raw | ConvertFrom-Json -AsHashtable
    Validate-ImportMap $importMap
    
    # Convert the import map to our structure, adding ID properties
    $notebookStructure = $importMap | ForEach-Object {
        @{
            OneNoteSectionGroupName = $_.OneNoteSectionGroupName
            OneNoteSectionGroupId = $null
            OneNoteSections = $_.OneNoteSections | ForEach-Object {
                @{
                    OneNoteSectionName = $_.OneNoteSectionName
                    OneNoteSectionId = $null
                    ThinkeryTags = $_.ThinkeryTags
                }
            }
        }
    }
    
    # Add a default section group and section for uncategorized content
    $notebookStructure += @{
        OneNoteSectionGroupName = $DEFAULT_GROUP_NAME
        OneNoteSectionGroupId = $null
        OneNoteSections = @(
            @{
                OneNoteSectionName = $DEFAULT_SECTION_NAME
                OneNoteSectionId = $null
                ThinkeryTags = @()
            }
        )
    }
    
    Write-Host "Import map loaded successfully with $($notebookStructure.Count) section groups."
}
catch {
    Write-Error "Error loading or validating import map: $_"
    exit 1
}

# Create section groups and sections based on our structure
foreach ($group in $notebookStructure) {
    # Create section group and store ID
    $group.OneNoteSectionGroupId = Create-SectionGroup -NotebookId $notebookId -Name $group.OneNoteSectionGroupName
    Write-Host "  Section Group: $($group.OneNoteSectionGroupName)"
    
    foreach ($section in $group.OneNoteSections) {
        # Create section and store ID
        $section.OneNoteSectionId = Create-Section -SectionGroupId $group.OneNoteSectionGroupId -Name $section.OneNoteSectionName
        Write-Host "    Section: $($section.OneNoteSectionName)"
    }
}

# Find the most appropriate section for a set of tags
Function Find-BestMatchSection($Tags) {
    # Track the best match
    $bestMatchGroup = $null
    $bestMatchSection = $null
    $bestMatchCount = -1
    $bestMatchPercentage = 0
    
    # Find the default group and section for uncategorized content
    $defaultGroup = $notebookStructure | Where-Object { $_.OneNoteSectionGroupName -eq $DEFAULT_GROUP_NAME }
    $defaultSection = $defaultGroup.OneNoteSections | Where-Object { $_.OneNoteSectionName -eq $DEFAULT_SECTION_NAME }
    
    # Go through each group and section to find the best tag match
    foreach ($group in $notebookStructure) {
        # Skip the default group during matching
        if ($group.OneNoteSectionGroupName -eq $DEFAULT_GROUP_NAME) {
            continue
        }
        
        foreach ($section in $group.OneNoteSections) {
            # Skip sections with no tags defined
            if ($section.ThinkeryTags.Count -eq 0) {
                continue
            }
            
            # Count how many tags match
            $matchCount = 0
            foreach ($tag in $Tags) {
                if ($section.ThinkeryTags -contains $tag) {
                    $matchCount++
                }
            }
            
            # Skip if no matches
            if ($matchCount -eq 0) {
                continue
            }
            
            # Calculate match quality:
            # 1. How many tags from the note match this section
            $matchRatio = [math]::Min(1.0, $matchCount / [math]::Max(1, $Tags.Count))
            
            # 2. How many section tags are matched (specificity)
            $specificityRatio = [math]::Min(1.0, $matchCount / [math]::Max(1, $section.ThinkeryTags.Count))
            
            # Combined score (emphasizes specificity slightly more)
            $matchPercentage = ($matchRatio * 0.4) + ($specificityRatio * 0.6)
            
            # Update best match if this is better
            if ($matchCount -gt $bestMatchCount || 
               ($matchCount -eq $bestMatchCount -and $matchPercentage -gt $bestMatchPercentage)) {
                $bestMatchGroup = $group
                $bestMatchSection = $section
                $bestMatchCount = $matchCount
                $bestMatchPercentage = $matchPercentage
            }
        }
    }
    
    # If no matches found, use the default uncategorized section
    if ($bestMatchCount -eq -1) {
        return @{ Group = $defaultGroup; Section = $defaultSection }
    }
    
    return @{ Group = $bestMatchGroup; Section = $bestMatchSection }
}

Write-Host "`nImporting pages..."

$agg = @{}   # sectionId|title|tags => [html fragments]

$json = Get-Content $JsonPath -Raw | ConvertFrom-Json
foreach ($n in $json) {
    $tags = ($n.tags -split "\s+") | Where-Object { $_ }
    
    # Find the best matching section based on tags
    $match = Find-BestMatchSection -Tags $tags
    $group = $match.Group
    $section = $match.Section
    
    if (!$group -or !$section) { 
        Write-Warning "No matching section found for note: $($n.title). Skipping."
        continue 
    }
    
    $secId = $section.OneNoteSectionId
    $groupName = $group.OneNoteSectionGroupName
    $sectionName = $section.OneNoteSectionName

    $created = [DateTime]::Parse($n.date).ToString("o")
    $title   = $n.title
    $content = $n.html
    $noteLen = $content.Length

    if ($noteLen -lt $TinyNoteThreshold) {
        # Sort tags alphabetically for consistent grouping
        $sortedTags = $tags | Sort-Object
        
        # Create tag string for the page title
        $tagString = if ($sortedTags.Count -gt 0) { 
            $sortedTags -join ", " 
        } else { 
            "untagged" 
        }
        
        # Create a descriptive page title including section name and tags
        $pageTitle = "Small notes - $sectionName - $tagString"
        
        # Create a key that includes section ID and tag string
        $key = "$secId|$tagString"
        
        # Initialize array if this is a new key
        if (!$agg.ContainsKey($key)) { 
            $agg[$key] = @{
                "title" = $pageTitle
                "fragments" = @()
            }
        }
        
        # Add this note to the appropriate aggregation group
        $agg[$key].fragments += "<h3>$title</h3><p>$content</p>"
        
        # Log the small note being aggregated
        $tagsString = if ($tags.Count -gt 0) { "'$($tags -join "', '")'" } else { "(no tags)" }
        Write-Host "  + Small note: '$title' → Aggregating to '$pageTitle' (in $groupName/$sectionName) [Tags: $tagsString]"
    } else {
        $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>$title</title>
    <meta name="created" content="$created"/>
</head>
<body>
$content
</body>
</html>
"@
        Post-Page -SectionId $secId -Html $html
        
        # Enhanced logging with full details
        $tagsString = if ($tags.Count -gt 0) { "'$($tags -join "', '")'" } else { "(no tags)" }
        Write-Host "  + Large page: '$title' → $groupName/$sectionName [Tags: $tagsString]"
    }
}

# Create the aggregated pages
foreach ($k in $agg.Keys) {
    $parts = $k -split '\|'
    $secId = $parts[0]
    $pageTitle = $agg[$k].title
    $body = ($agg[$k].fragments -join "`n")
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>$pageTitle</title>
    <meta name="created" content="$(Get-Date -Format o)"/>
</head>
<body>
$body
</body>
</html>
"@
    Post-Page -SectionId $secId -Html $html
    
    # Enhanced logging for aggregated pages
    $count = ($agg[$k].fragments.Count)
    Write-Host "  + Aggregated page: '$pageTitle' with $count notes"
}

Write-Host "`nImport complete!"

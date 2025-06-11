<#
.SYNOPSIS
    Import Thinkery JSON export into a new OneNote notebook via Microsoft Graph.

.DESCRIPTION
    1. Creates a new notebook (default name "Thinkery Import", override with -NotebookName).
    2. Builds section groups and sections according to the mapping specified in the import map.
    3. Parses the Thinkery JSON export, creating:
       * One page per "large" note (≥ 140 characters by default).
       * Aggregated pages for tiny notes grouped by tags.
    4. Uses interactive authentication with Microsoft Graph API.

.PARAMETER JsonPath
    Path to the Thinkery JSON export file.

.PARAMETER NotebookName
    Display name for the new notebook. Default: "Thinkery Import"

.PARAMETER ImportMapPath
    Path to the JSON file defining OneNote structure and tag mappings.
    Default: "./sample-import-maps/heywills-import-map.json"

.PARAMETER TinyNoteThreshold
    Character count threshold below which notes are considered "tiny" and will be aggregated.
    Default: 140

.PARAMETER LogPath
    Path where log files will be stored. Default: "./logs"

.PARAMETER DryRun
    If specified, the script will not make any changes, only report what it would do.

.EXAMPLE
    # Run with interactive authentication:
    .\Import-ThinkeryExportIntoOneNote.ps1 `
        -JsonPath ".\import-files\thinkery-export.json" `
        -ImportMapPath ".\sample-import-maps\my-import-map.json"

.EXAMPLE
    # Dry run to test without making changes:
    .\Import-ThinkeryExportIntoOneNote.ps1 `
        -JsonPath ".\import-files\thinkery-export.json" `
        -ImportMapPath ".\sample-import-maps\my-import-map.json" `
        -DryRun
#>

param(
    [Parameter(Mandatory = $true)][string]$JsonPath,
    [string]$NotebookName = "Thinkery Import",
    [Parameter(Mandatory = $true)][string]$ImportMapPath,
    [int]$TinyNoteThreshold = 140,
    [string]$LogPath = ".\\logs",
    [switch]$DryRun = $false
)

$ErrorActionPreference = "Break"
$graphApi = "https://graph.microsoft.com/v1.0"

# Setup logging
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$logFileName = "thinkery-import_$timestamp.log"

# Ensure log directory exists
if (-not (Test-Path $LogPath)) {
    try {
        New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
        Write-Host "Created log directory: $LogPath"
    }
    catch {
        Write-Warning "Could not create log directory: $LogPath. Logging to current directory instead."
        $LogPath = "."
    }
}

$logFile = Join-Path -Path $LogPath -ChildPath $logFileName

# Log function to write to both console and log file
Function Write-Log {
    param (
        [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Write to log file
    Add-Content -Path $logFile -Value $logMessage
    
    # Also write to console with color based on level
    switch ($Level) {
        "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
        "ERROR"   { Write-Host $logMessage -ForegroundColor Red }
        "SUCCESS" { Write-Host $logMessage -ForegroundColor Green }
        default   { Write-Host $logMessage }
    }
}

Write-Log "Starting Thinkery to OneNote import process" "INFO"
Write-Log "Log file: $logFile" "INFO"
Write-Log "Parameters:" "INFO"
Write-Log "  JsonPath: $JsonPath" "INFO"
Write-Log "  NotebookName: $NotebookName" "INFO"
Write-Log "  ImportMapPath: $ImportMapPath" "INFO"
Write-Log "  TinyNoteThreshold: $TinyNoteThreshold" "INFO"
Write-Log "  DryRun: $DryRun" "INFO"

# Interactive authentication function
Function Get-InteractiveAccessToken {
    Write-Log "Starting interactive authentication..." "INFO"
    
    # Check if MSAL.PS module is installed
    if (-not (Get-Module -ListAvailable -Name MSAL.PS)) {
        Write-Log "MSAL.PS module not found. Installing..." "INFO"
        try {
            Install-Module -Name MSAL.PS -Scope CurrentUser -Force -ErrorAction Stop
        }
        catch {
            Write-Log "Failed to install MSAL.PS module: $_" "ERROR"
            Write-Log "Please install the MSAL.PS module manually: Install-Module -Name MSAL.PS -Scope CurrentUser -Force" "ERROR"
            throw "MSAL.PS module is required for authentication"
        }
    }
    
    # Import the module
    Import-Module MSAL.PS
    
    $clientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e" # Microsoft Graph Explorer client ID
    $tenantId = "common"                               # Use 'common' for any tenant
    $redirectUri = "https://login.microsoftonline.com/common/oauth2/nativeclient"
    $scopes = @("https://graph.microsoft.com/Notes.ReadWrite")
    
    try {
        # Acquire token interactively
        Write-Log "Launching interactive login window..." "INFO"
        $authResult = Get-MsalToken -ClientId $clientId -TenantId $tenantId -RedirectUri $redirectUri `
                                  -Scopes $scopes -Interactive
        
        Write-Log "Authentication successful! Token acquired." "SUCCESS"
        return $authResult.AccessToken
    }
    catch {
        Write-Log "Failed to acquire access token: $_" "ERROR"
        throw "Authentication failed. $_"
    }
}

# Get access token if not in dry run mode
if ($DryRun) {
    Write-Log "Dry run mode - skipping authentication" "INFO"
    $AccessToken = "dry-run-token"
} else {
    $AccessToken = Get-InteractiveAccessToken
}

Function Invoke-GraphPost($Uri, $BodyObj) {
    $json = $BodyObj | ConvertTo-Json -Depth 6
    try {
        # For debugging
        Write-Debug "Sending request to $Uri with body: $json"
        
        if ($DryRun) {
            Write-Log "[DRY RUN] Would send request to $Uri" "INFO"
            return [PSCustomObject]@{ id = "dry-run-id-$(Get-Random)" }
        }
        
        Write-Log "Sending request to $Uri" "INFO"
        $response = Invoke-RestMethod -Method Post -Uri $Uri `
            -Headers @{ Authorization = "Bearer $AccessToken"; "Content-Type" = "application/json" } `
            -Body $json -ErrorVariable responseError
        return $response
    } catch {
        Write-Log "Graph API Error: $_" "ERROR"
        Write-Log "Request body: $json" "ERROR"
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
    Write-Log "Creating notebook '$Name' (sanitized as '$sanitizedName')..." "INFO"
    $nb = Invoke-GraphPost "$graphApi/me/onenote/notebooks" @{ displayName = $sanitizedName }
    return $nb.id
}

Function Create-SectionGroup {
    param([string]$NotebookId, [string]$Name)
    $sanitizedName = Sanitize-Name -Name $Name
    Write-Log "Creating section group '$Name' (sanitized as '$sanitizedName')..." "INFO"
    $sg = Invoke-GraphPost "$graphApi/me/onenote/notebooks/$NotebookId/sectionGroups" @{ displayName = $sanitizedName }
    return $sg.id
}

Function Create-Section {
    param([string]$SectionGroupId, [string]$Name)
    $sanitizedName = Sanitize-Name -Name $Name
    Write-Log "Creating section '$Name' (sanitized as '$sanitizedName')..." "INFO"
    $sec = Invoke-GraphPost "$graphApi/me/onenote/sectionGroups/$SectionGroupId/sections" @{ displayName = $sanitizedName }
    return $sec.id
}

Function Post-Page {
    param([string]$SectionId, [string]$Html)
    try {
        if ($DryRun) {
            Write-Log "[DRY RUN] Would post page to section $SectionId" "INFO"
            return
        }
        
        # Send the HTML content directly in the request body
        Write-Log "Posting page to section $SectionId" "INFO"
        $ret = Invoke-RestMethod -Method Post -Uri "$graphApi/me/onenote/sections/$SectionId/pages" `
            -Headers @{ 
                "Authorization" = "Bearer $AccessToken"
                "Content-Type" = "text/html; charset=utf-8"
            } `
            -Body $Html
    }
    catch {
        Write-Log "Error posting page: $_" "ERROR"
        Break
    }
}

Function Is-Checkbox {
    param([object]$Content)
    return $Content -is [bool]
}

Function Get-TodoTagAttribute {
    param([bool]$IsChecked)
    
    # Return the appropriate data-tag attribute for OneNote ToDo items
    if ($IsChecked) {
        return 'data-tag="todo:completed"'
    } else {
        return 'data-tag="to-do"'
    }
}

Function Create-OneNotePage {
    param(
        [string]$SectionId,
        [string]$Title,
        [object]$Content,
        [DateTime]$Created,
        [string]$GroupName,
        [string]$SectionName,
        [array]$Tags,
        [object]$Url = $null
    )
    
    $tagList = Get-SortedTagString -Tags $Tags
    
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>$Title</title>
    <meta name="created" content="$($Created.ToString('o'))"/>
</head>
<body>
<p>Created: $($Created.ToLocalTime().ToString("g"))</p>
$(Format-UrlLink -Url $Url)
<p>Tags: $tagList</p>
$(if (Is-Checkbox -Content $Content) {
    $todoAttr = Get-TodoTagAttribute -IsChecked $Content
    "<p $todoAttr>$Title</p>"
} else {
    $Content
})
</body>
</html>
"@
    Post-Page -SectionId $SectionId -Html $html
    
    # Enhanced logging with full details
    $tagsString = if ($Tags.Count -gt 0) { "'$($Tags -join "', '")'" } else { "(no tags)" }
    Write-Log "  + Large page: '$Title' → $GroupName/$SectionName [Tags: $tagsString]" "INFO"
    
    return $true
}

Function Create-OneNotePageWithTinyNotes {
    param(
        [string]$SectionId,
        [string]$PageTitle,
        [array]$Notes
    )
    
    # Build HTML body from note objects for aggregated pages with multiple notes
    $bodyFragments = $Notes | ForEach-Object {
        $createdDisplay = $_.created.ToLocalTime().ToString("g")
        $urlHtml = Format-UrlLink -Url $_.url
        
        if (Is-Checkbox -Content $_.content) {
            $todoAttr = Get-TodoTagAttribute -IsChecked $_.content
            "<h3 $todoAttr>$($_.title)</h3><p class='note-date'>Created: $createdDisplay</p>$urlHtml"
        } else {
            "<h3>$($_.title)</h3><p class='note-date'>Created: $createdDisplay</p>$urlHtml<p>$($_.content)</p>"
        }
    }
    $body = $bodyFragments -join "`n"
    
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>$PageTitle</title>
    <meta name="created" content="$(Get-Date -Format o)"/>
    <style>
        .note-date {
            color: #666;
            font-size: 0.9em;
            margin-top: -10px;
            font-style: italic;
        }
    </style>
</head>
<body>
$body
</body>
</html>
"@
    Post-Page -SectionId $SectionId -Html $html
    
    # Enhanced logging for aggregated pages
    $count = $Notes.Count
    Write-Log "  + Aggregated page: '$PageTitle' with $count notes" "SUCCESS"
}

Function Format-UrlLink {
    param([object]$Url)
    
    if ($Url -and $Url -isnot [bool] -and ![string]::IsNullOrEmpty($Url)) {
        return "<p><a href=`"$Url`">$Url</a></p>"
    }
    return ""
}

Function Get-SortedTagString {
    param ([array]$Tags)
    
    # Sort tags alphabetically for consistent grouping
    $sortedTags = $Tags | Sort-Object
    
    # Create tag string for the page title
    if ($sortedTags.Count -gt 0) { 
        return $sortedTags -join ", " 
    } else { 
        return "untagged" 
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
    Write-Log "[DRY RUN] Would create notebook '$NotebookName'" "INFO"
    $notebookId = "dry-run-notebook-id"
} else {
    $notebookId = Create-Notebook -Name $NotebookName
}
Write-Log "Notebook created with id $notebookId" "SUCCESS"

# 2. Load and validate import map
try {
    Write-Log "Loading import map from $ImportMapPath..." "INFO"
    
    if (-not (Test-Path $ImportMapPath)) {
        Write-Log "Import map file not found: $ImportMapPath" "ERROR"
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
    
    Write-Log "Import map loaded successfully with $($notebookStructure.Count) section groups." "SUCCESS"
}
catch {
    Write-Log "Error loading or validating import map: $_" "ERROR"
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

Write-Log "`nImporting pages..." "INFO"

$agg = @{}   # sectionId|title|tags => [html fragments]

$json = Get-Content $JsonPath -Raw | ConvertFrom-Json
$totalNotes = $json.Length
Write-Log "Found $totalNotes notes to import" "INFO"

$largeNoteCount = 0
$smallNoteCount = 0

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

    $created = [DateTime]::Parse($n.date) 
    $title   = $n.title
    $content = $n.html
    $noteLen = $content.Length
    $url     = $n.url

    if ($noteLen -lt $TinyNoteThreshold) {
        # Create tag string for the page title
        $tagString = Get-SortedTagString -Tags $tags
        
        # Create a descriptive page title including section name and tags
        $pageTitle = "Small notes - $sectionName - $tagString"
        
        # Create a key that includes section ID and tag string
        $key = "$secId|$tagString"
        
        # Initialize array if this is a new key
        if (!$agg.ContainsKey($key)) { 
            $agg[$key] = @{
                "title" = $pageTitle
                "sectionId" = $secId
                "notes" = @()
            }
        }
        
        # Add this note as a complete object to the appropriate aggregation group
        $agg[$key].notes += @{
            "title" = $title
            "content" = $content
            "created" = $created
            "groupName" = $groupName
            "sectionName" = $sectionName
            "tags" = $tags
            "url" = $url
            "noteLen" = $noteLen
        }
        
        # Log the small note being aggregated
        $tagsString = if ($tags.Count -gt 0) { "'$($tags -join "', '")'" } else { "(no tags)" }
        Write-Log "  + Small note: '$title' → Aggregating to '$pageTitle' (in $groupName/$sectionName) [Tags: $tagsString]" "INFO"
        $smallNoteCount++
    } else {
        Create-OneNotePage -SectionId $secId -Title $title -Content $content -Created $created `
                          -GroupName $groupName -SectionName $sectionName -Tags $tags -Url $url
        $largeNoteCount++
    }
}

$aggregatedPageCount = $agg.Keys.Count

# Create the aggregated pages
foreach ($k in $agg.Keys) {
    $parts = $k -split '\|'
    $secId = $parts[0]
    $pageTitle = $agg[$k].title
    $notes = $agg[$k].notes
    
    # If there's only one note in this group, create it as a regular page instead of aggregating
    if ($notes.Count -eq 1) {
        $note = $notes[0]
        Create-OneNotePage -SectionId $secId -Title $note.title -Content $note.content -Created $note.created `
                         -GroupName $note.groupName -SectionName $note.sectionName -Tags $note.tags -Url $note.url
        
        # Update the logging message to indicate this was handled as a single page
        Write-Log "  + Small note promoted to full page: '$($note.title)'" "SUCCESS"
    }
    else {
        Create-OneNotePageWithTinyNotes -SectionId $secId -PageTitle $pageTitle -Notes $notes
    }
}

Write-Log "`nImport summary:" "SUCCESS"
Write-Log "  Total notes processed: $totalNotes" "SUCCESS" 
Write-Log "  Large notes (individual pages): $largeNoteCount" "SUCCESS"
Write-Log "  Small notes (aggregated): $smallNoteCount" "SUCCESS"
Write-Log "  Aggregated pages created: $aggregatedPageCount" "SUCCESS"
Write-Log "`nImport complete!" "SUCCESS"

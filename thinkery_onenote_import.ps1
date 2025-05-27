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

# 1. Notebook
if ($DryRun) {
    Write-Host "[DRY RUN] Would create notebook '$NotebookName'" -ForegroundColor Yellow
    $notebookId = "dry-run-notebook-id"
} else {
    $notebookId = Create-Notebook -Name $NotebookName
}
Write-Host "Notebook created with id $notebookId"

# 2. Section Groups & Sections
$notebookStructure = @(
    @{
        OneNoteSectionGroupName = "Bible"
        OneNoteSectionGroupId = $null
        OneNoteSections = @(
            @{
                OneNoteSectionName = "Memory Verses"
                OneNoteSectionId = $null
                ThinkeryTags = @("memory_verse")
            },
            @{
                OneNoteSectionName = "Study Notes"
                OneNoteSectionId = $null
                ThinkeryTags = @()  # Default for Bible group
            }
        )
    },
    @{
        OneNoteSectionGroupName = "Homeschool"
        OneNoteSectionGroupId = $null
        OneNoteSections = @(
            @{
                OneNoteSectionName = "Task Lists"
                OneNoteSectionId = $null
                ThinkeryTags = @("todo")
            },
            @{
                OneNoteSectionName = "Courses & Resources"
                OneNoteSectionId = $null
                ThinkeryTags = @()  # Default for Homeschool group
            }
        )
    },
    @{
        OneNoteSectionGroupName = "Gift Ideas"
        OneNoteSectionGroupId = $null
        OneNoteSections = @(
            @{
                OneNoteSectionName = "Amy"
                OneNoteSectionId = $null
                ThinkeryTags = @("amy")
            },
            @{
                OneNoteSectionName = "Jessica"
                OneNoteSectionId = $null
                ThinkeryTags = @("jessica")
            },
            @{
                OneNoteSectionName = "Joseph"
                OneNoteSectionId = $null
                ThinkeryTags = @("joseph")
            },
            @{
                OneNoteSectionName = "General"
                OneNoteSectionId = $null
                ThinkeryTags = @()  # Default for Gift Ideas group
            }
        )
    },
    @{
        OneNoteSectionGroupName = "Wish Lists"
        OneNoteSectionGroupId = $null
        OneNoteSections = @(
            @{
                OneNoteSectionName = "Mike"
                OneNoteSectionId = $null
                ThinkeryTags = @("mike", "me")
            }
        )
    },
    @{
        OneNoteSectionGroupName = "Health"
        OneNoteSectionGroupId = $null
        OneNoteSections = @(
            @{
                OneNoteSectionName = "Family Health"
                OneNoteSectionId = $null
                ThinkeryTags = @("family")
            },
            @{
                OneNoteSectionName = "Personal Trackers"
                OneNoteSectionId = $null
                ThinkeryTags = @()  # Default for Health group
            }
        )
    },
    @{
        OneNoteSectionGroupName = "Home"
        OneNoteSectionGroupId = $null
        OneNoteSections = @(
            @{
                OneNoteSectionName = "Maintenance Log"
                OneNoteSectionId = $null
                ThinkeryTags = @("maintenance", "repair")
            },
            @{
                OneNoteSectionName = "Contractors & Quotes"
                OneNoteSectionId = $null
                ThinkeryTags = @("contractor", "quote")
            }
        )
    },
    @{
        OneNoteSectionGroupName = "Outdoor & Gear"
        OneNoteSectionGroupId = $null
        OneNoteSections = @(
            @{
                OneNoteSectionName = "Gear Lists"
                OneNoteSectionId = $null
                ThinkeryTags = @()  # Default for Outdoor & Gear group
            },
            @{
                OneNoteSectionName = "Trip Planning"
                OneNoteSectionId = $null
                ThinkeryTags = @("trip")
            }
        )
    },
    @{
        OneNoteSectionGroupName = "Pets"
        OneNoteSectionGroupId = $null
        OneNoteSections = @(
            @{
                OneNoteSectionName = "Health History"
                OneNoteSectionId = $null
                ThinkeryTags = @("health", "vet")
            },
            @{
                OneNoteSectionName = "Food & Supplies"
                OneNoteSectionId = $null
                ThinkeryTags = @("food", "supplies")
            }
        )
    },
    @{
        OneNoteSectionGroupName = "Hunting & Shooting"
        OneNoteSectionGroupId = $null
        OneNoteSections = @(
            @{
                OneNoteSectionName = "Turkey Tips"
                OneNoteSectionId = $null
                ThinkeryTags = @("turkey", "turkey_tips")
            },
            @{
                OneNoteSectionName = "Deer & Elk"
                OneNoteSectionId = $null
                ThinkeryTags = @("deer", "elk")
            },
            @{
                OneNoteSectionName = "Equipment"
                OneNoteSectionId = $null
                ThinkeryTags = @()  # Default for Hunting & Shooting group
            }
        )
    },
    @{
        OneNoteSectionGroupName = "Misc Notes"
        OneNoteSectionGroupId = $null
        OneNoteSections = @(
            @{
                OneNoteSectionName = "Tech Snippets"
                OneNoteSectionId = $null
                ThinkeryTags = @("tech", "tech_tip", "tips")
            },
            @{
                OneNoteSectionName = "Quotes"
                OneNoteSectionId = $null
                ThinkeryTags = @("quotes")
            },
            @{
                OneNoteSectionName = "Inbox"
                OneNoteSectionId = $null
                ThinkeryTags = @()  # Default for Misc Notes group
            }
        )
    },
    @{
        OneNoteSectionGroupName = "Book Trackers"
        OneNoteSectionGroupId = $null
        OneNoteSections = @(
            @{
                OneNoteSectionName = "Series"
                OneNoteSectionId = $null
                ThinkeryTags = @()  # Default for Book Trackers group
            }
        )
    },
    @{
        OneNoteSectionGroupName = "Blog / Writing Ideas"
        OneNoteSectionGroupId = $null
        OneNoteSections = @(
            @{
                OneNoteSectionName = "CMS Ideas"
                OneNoteSectionId = $null
                ThinkeryTags = @("cms", "kentico", "sitefinity")
            },
            @{
                OneNoteSectionName = "App Ideas"
                OneNoteSectionId = $null
                ThinkeryTags = @("app", "utility")
            },
            @{
                OneNoteSectionName = "Other Drafts"
                OneNoteSectionId = $null
                ThinkeryTags = @()  # Default for Blog / Writing Ideas group
            }
        )
    }
)

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
    
    # Default to Misc Notes/Inbox if no match is found
    $defaultGroup = $notebookStructure | Where-Object { $_.OneNoteSectionGroupName -eq "Misc Notes" }
    $defaultSection = $defaultGroup.OneNoteSections | Where-Object { $_.OneNoteSectionName -eq "Inbox" }
    
    # Go through each group and section to find the best tag match
    foreach ($group in $notebookStructure) {
        foreach ($section in $group.OneNoteSections) {
            # Count how many tags match
            $matchCount = 0
            foreach ($tag in $Tags) {
                if ($section.ThinkeryTags -contains $tag) {
                    $matchCount++
                }
            }
            
            # Update best match if this is better
            if ($matchCount -gt $bestMatchCount) {
                $bestMatchGroup = $group
                $bestMatchSection = $section
                $bestMatchCount = $matchCount
            }
        }
    }
    
    # If no matches found, use default section or first section in group
    if ($bestMatchCount -eq 0) {
        # Try to find a default section in the first matching group
        foreach ($group in $notebookStructure) {
            $defaultGroupSection = $group.OneNoteSections | Where-Object { $_.ThinkeryTags.Count -eq 0 }
            if ($defaultGroupSection) {
                return @{ Group = $group; Section = $defaultGroupSection }
            }
        }
        
        # If still no match, use the global default
        return @{ Group = $defaultGroup; Section = $defaultSection }
    }
    
    return @{ Group = $bestMatchGroup; Section = $bestMatchSection }
}

Write-Host "`nImporting pages..."

$tinyThreshold = 300
$agg = @{}   # sectionId|title => [html fragments]

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

    if ($noteLen -lt $tinyThreshold) {
        $pageTitle = "Quick Notes - $sectionName"
        $key = "$secId|$pageTitle"
        if (!$agg.ContainsKey($key)) { $agg[$key] = @() }
        $agg[$key] += "<h3>$title</h3><p>$content</p>"
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
        Write-Host "  + Large page: $title (in $groupName/$sectionName)"
    }
}

foreach ($k in $agg.Keys) {
    $parts = $k -split '\|'
    $secId = $parts[0]
    $pageTitle = $parts[1]
    $body = ($agg[$k] -join "`n")
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
    Write-Host "  + Aggregated page: $pageTitle"
}

Write-Host "`nImport complete!"

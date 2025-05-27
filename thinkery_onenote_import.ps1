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
    [string]$JsonPath = ".\\thinkery-tiriansdoor.json",
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
$sectionGroups = @("Bible","Homeschool","Gift Ideas","Wish Lists","Health","Home","Outdoor & Gear","Pets","Hunting & Shooting","Misc Notes","Book Trackers","Blog / Writing Ideas")
$sectionsPerGroup = @{
    "Bible" = @("Memory Verses","Study Notes");
    "Homeschool" = @("Task Lists","Courses & Resources");
    "Gift Ideas" = @("Amy","Jessica","Joseph","General");
    "Wish Lists" = @("Mike");
    "Health" = @("Family Health","Personal Trackers");
    "Home" = @("Maintenance Log","Contractors & Quotes");
    "Outdoor & Gear" = @("Gear Lists","Trip Planning");
    "Pets" = @("Health History","Food & Supplies");
    "Hunting & Shooting" = @("Turkey Tips","Deer & Elk","Equipment");
    "Misc Notes" = @("Tech Snippets","Quotes","Inbox");
    "Book Trackers" = @("Series");
    "Blog / Writing Ideas" = @("CMS Ideas","App Ideas","Other Drafts");
}

$sgIds   = @{}
$secIds  = @{}

foreach ($sg in $sectionGroups) {
    $sgIds[$sg] = Create-SectionGroup -NotebookId $notebookId -Name $sg
    Write-Host "  Section Group: $sg"
    foreach ($sec in $sectionsPerGroup[$sg]) {
        $secIds["$sg|$sec"] = Create-Section -SectionGroupId $sgIds[$sg] -Name $sec
        Write-Host "    Section: $sec"
    }
}

# 3. Helper functions for note routing
Function Get-FolderTag($Tags) {
    if ($null -eq $Tags -or $Tags.Count -eq 0) {
        return "misc_notes_folder"
    }
    
    $folderTags = $Tags | Where-Object { $_ -like "*_folder" }
    if ($null -eq $folderTags -or $folderTags.Count -eq 0) {
        return "misc_notes_folder"
    }
    
    # Check if $folderTags is a string or an array
    if ($folderTags -is [string]) {
        return $folderTags
    } else {
        return $folderTags[0]
    }
}

Function MapFolderToGroup($Folder) {
    switch ($Folder) {
        "bible_folder"       { return "Bible" }
        "homeschool_folder"  { return "Homeschool" }
        "gift_ideas_folder"  { return "Gift Ideas" }
        "wish_list_folder"   { return "Wish Lists" }
        "misc_notes_folder"  { return "Misc Notes" }
        "outdoor_folder"    { return "Outdoor & Gear" }
        "outdoors_folder"   { return "Outdoor & Gear" }
        "prayers_folder"  { return "Bible" }
        default              { return "Misc Notes" }
    }
}

# Function MapToSection($Group,$Tags) {
#     switch ($Group) {
#         "Bible"              { if ($Tags -contains "memory_verse") { return "Memory Verses" } else { return "Study Notes" } }
#         "Homeschool"         { if ($Tags -contains "todo") { return "Task Lists" } else { return "Courses & Resources" } }
#         "Gift Ideas"         { $p = $Tags | Where-Object { $_ -in "mike","me" }; if ($p){ return ($p[0] -replace "^.").ToUpper()+$p[0].Substring(1) } else { return "General" } }
#         "Wish Lists"         { $p = $Tags | Where-Object { $_ -in "amy","jessica","joseph"}; if ($p){ return ($p[0] -replace "^.").ToUpper()+$p[0].Substring(1) } else { return "General" } }
#         default              { return $sectionsPerGroup[$Group][0] }
#     }
# }

Function MapToSection($Group, $Tags) {
    # Add breakpoint if any element in $Tags is a char
    foreach ($tag in $Tags) {
        if ($tag -is [char]) {
            Write-Host "Breakpoint: Found a char in Tags: '$tag'" -ForegroundColor Red
            Break
        }
    }
    
    switch ($Group) {
        "Bible" {
            if ($Tags -contains "memory_verse") { return "Memory Verses" }
            else { return "Study Notes" }
        }
        "Homeschool" {
            if ($Tags -contains "todo") { return "Task Lists" }
            else { return "Courses & Resources" }
        }
        "Gift Ideas" {
            $p = $Tags | Where-Object { $_ -in "mike","me","jessica","amy","joseph" }
            if ($p) {
                return ($p[0].Substring(0,1).ToUpper() + $p[0].Substring(1))
            } else {
                return "General"
            }
        }
        "Wish Lists" {
            $p = $Tags | Where-Object { $_ -in "mike","amy","jessica","joseph" }
            if ($p) {
                return ($p[0].Substring(0,1).ToUpper() + $p[0].Substring(1))
            } else {
                return "General"
            }
        }
        "Misc Notes" {
            if ($Tags -contains "quotes") { return "Quotes" }
            elseif ($Tags -contains "tech" -or $Tags -contains "tech_tip" -or $Tags -contains "tips") {
                return "Tech Snippets"
            } else {
                return "Inbox"
            }
        }
        "Hunting & Shooting" {
            if ($Tags -contains "turkey" -or $Tags -contains "turkey_tips") { return "Turkey Tips" }
            elseif ($Tags -contains "deer" -or $Tags -contains "elk") { return "Deer & Elk" }
            else { return "Equipment" }
        }
        "Outdoor & Gear" {
            if ($Tags -contains "trip") { return "Trip Planning" }
            else { return "Gear Lists" }
        }
        "Blog / Writing Ideas" {
            if ($Tags -contains "cms" -or $Tags -contains "kentico" -or $Tags -contains "sitefinity") {
                return "CMS Ideas"
            } elseif ($Tags -contains "app" -or $Tags -contains "utility") {
                return "App Ideas"
            } else {
                return "Other Drafts"
            }
        }
        default {
            return $sectionsPerGroup[$Group][0]
        }
    }
}


Write-Host "`nImporting pages..."

$tinyThreshold = 300
$agg = @{}   # sectionId|title => [html fragments]

$json = Get-Content $JsonPath -Raw | ConvertFrom-Json
foreach ($n in $json) {
    $tags = ($n.tags -split "\s+") | Where-Object { $_ }
    $folder = Get-FolderTag $tags
    $group  = MapFolderToGroup $folder
    $section = MapToSection $group $tags
    $secId = $secIds["$group|$section"]

    if (!$secId) { 
        Write-Warning "No section found for group '$group' and section '$section'. Skipping note: $($n.title)"
        continue 
    }

    $created = [DateTime]::Parse($n.date).ToString("o")
    $title   = $n.title
    $content = $n.html
    $noteLen = $content.Length

    if ($noteLen -lt $tinyThreshold) {
        $pageTitle = "Quick Notes - $section"
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
        Write-Host "  + Large page: $title"
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

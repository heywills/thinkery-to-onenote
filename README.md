# Thinkery to OneNote Import Tool

A PowerShell script that imports Thinkery JSON exports into Microsoft OneNote using the Microsoft Graph API.

## Purpose

The `Import-ThinkeryExportIntoOneNote.ps1` script helps you migrate your notes from Thinkery to Microsoft OneNote. It:

1. Creates a new OneNote notebook with customizable structure
2. Organizes notes into section groups and sections based on a mapping configuration
3. Categorizes notes using tags to find the most appropriate location
4. Creates individual pages for larger notes
5. Aggregates smaller notes into topic-based pages

## How to Use

### Prerequisites

- PowerShell 5.1 or higher
- A Microsoft account with access to OneNote
- A JSON export of your Thinkery notes

### Step 1: Prepare Your Import File

1. Export your notes from Thinkery in JSON format
2. Place your Thinkery JSON export in the `import-files` folder
3. Make note of the exact path to use with the `-JsonPath` parameter

### Step 2: Create Your Import Mapping Configuration

1. Create or customize an import map JSON file based on the structure below
2. Save your import map in the `sample-import-maps` folder or another location
3. Make note of the exact path to use with the `-ImportMapPath` parameter

### Step 3: Get an Access Token

1. Open [Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer)
2. Sign in with your Microsoft account
3. Set the permissions to Notes.ReadWrite:
   - Choose API Explorer from the left menu
   - Expand me->onenote->notebooks->post
   - Click "Modify permissions"
   - Click Consent next to "Notes.ReadWrite"
4. Click the "Access token" section and copy the token

### Step 4: Run the Script

```powershell
.\Import-ThinkeryExportIntoOneNote.ps1 `
    -AccessToken "your-access-token-here" `
    -JsonPath ".\import-files\your-thinkery-export.json" `
    -ImportMapPath ".\sample-import-maps\your-import-map.json"
```

### Optional Parameters

- `-NotebookName "Your Notebook Name"` - Customize the name of the OneNote notebook (default: "Thinkery Import")
- `-TinyNoteThreshold 140` - Character threshold for aggregating small notes (default: 140 characters)
- `-LogPath ".\custom-logs"` - Custom path for logs (default: ".\logs")
- `-DryRun` - Test the script without making changes to OneNote

## Import Map Structure

The import map is a JSON file that defines how your notes should be organized in OneNote. It specifies section groups, sections, and tag mappings.

### Example

```json
[
  {
    "OneNoteSectionGroupName": "Work",
    "OneNoteSections": [
      {
        "OneNoteSectionName": "Projects",
        "ThinkeryTags": ["project", "work_project", "deadline"]
      },
      {
        "OneNoteSectionName": "Meeting Notes",
        "ThinkeryTags": ["meeting", "call", "standup"]
      }
    ]
  },
  {
    "OneNoteSectionGroupName": "Personal",
    "OneNoteSections": [
      {
        "OneNoteSectionName": "Journal",
        "ThinkeryTags": ["journal", "diary", "reflection"]
      },
      {
        "OneNoteSectionName": "Ideas",
        "ThinkeryTags": ["idea", "inspiration", "concept"]
      }
    ]
  }
]
```

### Structure Details

1. **Top Level Array**: A list of section groups
   
2. **Section Group Object**:
   - `OneNoteSectionGroupName`: Name of the section group in OneNote
   - `OneNoteSections`: Array of section objects

3. **Section Object**:
   - `OneNoteSectionName`: Name of the section in OneNote
   - `ThinkeryTags`: Array of Thinkery tags that should be mapped to this section

### How Mapping Works

When importing notes, the script:

1. Examines tags on each Thinkery note
2. Finds the section with the best match based on the tags listed in `ThinkeryTags`
3. Places large notes as individual pages
4. Groups smaller notes into aggregated pages by section and tags

## Advanced Features

- **Note Aggregation**: Small notes (below the character threshold) with the same tags are grouped into pages titled "Small notes - [Section Name] - [tag1, tag2, ...]"
- **Tag Matching Algorithm**: Notes are placed based on the best match between their tags and the section's ThinkeryTags
- **Default Section**: Notes that don't match any tags will be placed in an "Uncategorized" section
- **Comprehensive Logging**: All operations are logged to both console and file for tracking progress and troubleshooting

## Troubleshooting

If you encounter issues:

1. Check the log files in the `logs` directory
2. Ensure your access token is valid (they expire relatively quickly)
3. Verify your import map file has the correct structure
4. Make sure your Thinkery export JSON is valid

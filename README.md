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

- PowerShell 7.1 or higher
- A Microsoft account with access to OneNote
- A JSON export of your Thinkery notes

### Step 1: Prepare Your Import File

1. Export your notes from Thinkery in JSON format
2. Place your Thinkery JSON export in the `import-files` folder
3. Make note of the exact path to use with the `-JsonPath` parameter

### Step 2: Create Your Import Mapping Configuration

1. Create an import map JSON file based on the structure below. Consider using ChatGPT to help you create the mapping configuration. A sample prompt is provided at the end of this readme.
2. Save your import map in the `sample-import-maps` folder or another location
3. Make note of the exact path to use with the `-ImportMapPath` parameter

### Step 3: Run the Script

The script uses interactive authentication to connect to Microsoft Graph API.

Make sure you run the script in PowerShell 7 (`pwsh.exe`) not PowerShell 5.1 (`powershell.exe`).

```powershell
.\Import-ThinkeryExportIntoOneNote.ps1 `
    -JsonPath ".\import-files\your-thinkery-export.json" `
    -ImportMapPath ".\sample-import-maps\your-import-map.json"
```

### Optional Parameters

- `-NotebookName "Your Notebook Name"` - Customize the name of the OneNote notebook (default: "Thinkery Import")
- `-TinyNoteThreshold 140` - Character threshold for aggregating small notes (default: 140 characters)
- `-LogPath ".\custom-logs"` - Custom path for logs (default: ".\logs")
- `-DryRun` - Test the script without making changes to OneNote (skips authentication)

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
2. Run the script with `-DryRun` parameter to test without making any changes
3. If authentication fails, restart your application and try again
4. Verify your import map file has the correct structure
5. Make sure your Thinkery export JSON is valid

## Sample ChatGPT prompt for creating a mapping file

I'm using an import script that maps tagged notes from a legacy note taking application to OneNote section groups, and sections.

The tool requires a json config file that defines the OneNote section groups and sections to create, and it requires mapping the sections to the tags from the legacy note taking app.

Here is a sample of the format of the json config file:

```json
[
  {
    "OneNoteSectionGroupName": "Section group 1",
    "OneNoteSections": [
      {
        "OneNoteSectionName": "Section 1a",
        "ThinkeryTags": ["legacy_tag_1", "legacy_tag_2", "legacy_tag_3"]  
      },
      {
        "OneNoteSectionName": "Section 1b",
        "ThinkeryTags": ["legacy_tag_4", "legacy_tag_5"]  
      }
    ]
  },
  {
    "OneNoteSectionGroupName": "Section group 2",
    "OneNoteSections": [
      {
        "OneNoteSectionName": "Section 2a",
        "ThinkeryTags": ["legacy_tag_6", "legacy_tag_7"]  
      },
      {
        "OneNoteSectionName": "Section 2b",
        "ThinkeryTags": ["legacy_tag_8", "legacy_tag_9"]  
      }
    ]
  }
]
```

Will you analyze the the JSON export, "thinkery-mikewills.json", from the legacy note application?

Please do the following:

- Recommend the hierarchy of OneNote section groups and sections to organize all the content in the attached JSON file, base the recommendations based on the tags and tag combinations.
- Please recognize that tags in the legacy JSON file are space-delimited and can have underscores (_) in their tag names.
- Please do a semantic analysis of the tag names to discovery their meaning and intent and try to create section groups and section names with semantic meaning? You may have to examine the note titles assigned to the tags to discovery meaning.
- Assign all the tags in the attached JSON file to one of the sections. You can combine multiple tags in one section if they are used together frequently.
- Create a configuration file in the above format that defines the recommended section groups, sections, and tag assignments.

Here's a starting point of the section group and section hierarchy reflecting section groups and sections that I know I want, but I'll need you to add more:

- Bible
  - Memory Verses
  - Study Notes
- Homeschool
  - Task Lists
  - Courses & Resources
- Gift Ideas
  - Amy
  - Jessica
  - Joseph
  - General
- Wish Lists
  - Mike
- Health
  - Family Health
  - Personal Trackers
- Home
  - Maintenance Log
  - Contractors & Quotes
- Outdoor & Gear
  - Gear Lists
  - Trip Planning
- Pets
  - Health History
  - Food & Supplies
- Hunting & Shooting
  - Turkey Tips
  - Deer & Elk
  - Equipment
- Places
  - Aeneas Valley
- Technology
  - Kentico
  - SharePoint
  - Code snippets
- Misc Notes
  - Tech Snippets
  - Quotes
  - Inbox
- Reading and literature
  - Series
  - Quotes
- Blog - Writing Ideas
  - CMS Ideas
  - App Ideas
  - Other Drafts

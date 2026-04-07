# Power Automate Flow: Search SharePoint Folders for "piv" Files and Log to Excel

## Overview
This guide explains how to create a Power Automate flow that:
1. Accesses SharePoint document libraries with nested folders
2. Recursively navigates to the deepest folders
3. Checks for files containing "piv" in the filename
4. Logs the folder path and Yes/No status to an Excel file on SharePoint

## Prerequisites
- SharePoint site with document library containing nested folders
- Excel file created on SharePoint with the following columns:
  - **FolderPath** (Text) - The path to the folder being checked
  - **HasPivFile** (Text) - "Yes" or "No" indicating if a "piv" file exists
  - **CheckedDate** (Date/Time) - When the check was performed

## Step-by-Step Implementation

### Step 1: Create a New Instant or Scheduled Flow
1. Go to [Power Automate](https://make.powerautomate.com)
2. Click **Create** > Choose either:
   - **Instant cloud flow** (manual trigger)
   - **Scheduled cloud flow** (runs automatically on a schedule)
3. Name your flow (e.g., "Search PIV Files in SharePoint")

### Step 2: Add Trigger
- For **Instant flow**: Use "Manually trigger a flow" trigger
- For **Scheduled flow**: Use "Recurrence" trigger and set your desired schedule

### Step 3: Initialize Variables
Add a **"Initialize variable"** action for each variable:

| Variable Name | Type | Value |
|--------------|------|-------|
| `siteAddress` | String | Your SharePoint site address (e.g., `https://yourcompany.sharepoint.com/sites/YourSite`) |
| `libraryName` | String | Document library name (e.g., `Shared Documents`) |
| `excelLocation` | String | Path to Excel file (e.g., `/Shared Documents/PIV-Results.xlsx`) |
| `table` | String | Excel table name (e.g., `Table1`) |

### Step 4: Get Root Folders
Add **"Get files (properties only)"** action:
- **Site Address**: Select your SharePoint site
- **Library Name**: Select your document library
- **Folder**: Leave blank or set to `/` for root
- **Limit scope**: Select "All items"

### Step 5: Apply to Each Root Item
Add **"Apply to each"** control:
- Select `value` from the previous "Get files" action

### Step 6: Check if Item is a Folder
Inside the loop, add a **Condition** control:
- **Condition**: `ContentType` contains `Folder` OR check if `{IsFolder}` equals `true`
  - Alternative: Check if `FileRef` contains `.aspx` (folders typically don't have file extensions)

### Step 7: If Yes (It's a Folder) - Get Subfolders
Inside the "If yes" branch:

1. Add **"Get files (properties only)"** action:
   - **Site Address**: Your SharePoint site
   - **Library Name**: Your document library
   - **Folder**: Select `FullPath` or `ServerRelativeUrl` from current item
   - **Limit scope**: All items

2. Add another **"Apply to each"** for the subfolders
   - Select `value` from the subfolder query

3. **Continue nesting** the logic recursively OR use a simpler approach:
   - For each folder found, check if it has subfolders
   - If no subfolders found, it's a "deepest" folder - proceed to check for PIV files

### Step 8: Simplified Approach - Check All Folders
A simpler approach is to check ALL folders (not just deepest):

1. **Get all files and folders** from the library with recursive query
2. **Filter** for items where `ContentType` = `Folder`
3. For **each folder**, check for files containing "piv"

### Step 9: Check for PIV Files in Each Folder
For each folder (deepest or all):

1. Add **"Get files (properties only)"** action:
   - **Site Address**: Your SharePoint site
   - **Library Name**: Your document library
   - **Folder**: Current folder path
   - **Filter Query**: `substringof('piv', Name) eq true` (filters for files with "piv" in name)
   - **Limit scope**: All items

2. Add **Condition** to check if any files were found:
   - Check if `length(body('Get_files_(properties_only)')?['value'])` is greater than 0

### Step 10: Add Row to Excel
Add **"Add a row into a table"** action (Excel Online Business connector):

- **Location**: Your SharePoint site
- **Document Library**: Your document library
- **File**: Path to your Excel file
- **Table**: Your table name

**Fields:**
- **FolderPath**: `FullPath` or `ServerRelativeUrl` from current folder
- **HasPivFile**: 
  - In "If yes" branch: `"Yes"`
  - In "If no" branch: `"No"`
- **CheckedDate**: `utcNow()` or `formatDateTime(utcNow(), 'yyyy-MM-dd HH:mm:ss')`

## Complete Flow Structure

```
Trigger (Manual/Scheduled)
    ↓
Initialize Variables
    ↓
Get Root Folders (Get files properties only)
    ↓
Apply to Each (root items)
    ↓
    Condition: Is Folder?
    ├─ Yes → Get Subfolders
    │         ↓
    │         Apply to Each (subfolders)
    │         ↓
    │         Condition: Is Folder?
    │         ├─ Yes → [Continue recursively OR process as deepest folder]
    │         └─ No → Skip (it's a file)
    │
    └─ No → Skip (it's a file, not a folder)
```

## Alternative: Using Send an HTTP Request to SharePoint
For more control, use HTTP requests:

1. Add **"Send an HTTP request to SharePoint"** action
2. Use SharePoint REST API:
   ```
   _api/web/GetFolderByServerRelativeUrl('/sites/YourSite/Shared Documents/FolderPath')/Files
   ```
3. Parse the JSON response to check for files with "piv" in the name

## Excel File Setup
Create an Excel file on SharePoint with a table:

1. Create a new Excel workbook
2. Add headers in row 1: `FolderPath`, `HasPivFile`, `CheckedDate`
3. Select the range and press **Ctrl+T** to create a table
4. Name the table (e.g., `PIVResults`)
5. Save the file to your SharePoint document library

## Testing the Flow
1. Save and test your flow
2. Check the Excel file to verify results are being added
3. Review flow run history for any errors

## Error Handling
Add **Configure run after** settings:
- If any action fails, you can send an email notification
- Add a **Scope** with "Configure run after" to handle failures gracefully

## Sample Flow JSON (for Import)
You can export/import flows. Here's a basic structure to adapt:

```json
{
  "trigger": "Recurrence",
  "actions": [
    {
      "type": "InitializeVariable",
      "inputs": {
        "variables": [
          {
            "name": "siteAddress",
            "type": "string",
            "value": "https://yourcompany.sharepoint.com/sites/YourSite"
          }
        ]
      }
    },
    {
      "type": "GetFiles",
      "inputs": {
        "siteAddress": "@variables('siteAddress')",
        "libraryName": "Shared Documents",
        "folderPath": "/",
        "limitScope": "AllItems"
      }
    }
  ]
}
```

## Tips
1. **Performance**: If you have many folders, consider batching or limiting the scope
2. **Permissions**: Ensure the flow has access to both the source folders and the Excel file
3. **File Naming**: The check is case-insensitive by default in SharePoint filters
4. **Duplicates**: If running on a schedule, consider clearing the Excel file first or appending with timestamps

## Final Notes
- Test with a small subset of folders first
- Monitor flow runs for any throttling issues with SharePoint
- Consider adding logging or notifications for completion status
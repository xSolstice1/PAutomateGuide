# Excel Template Setup for PIV File Search Results

## Excel File Structure

Create an Excel file on SharePoint with the following structure:

### Column Headers (Row 1)
| Column A | Column B | Column C |
|----------|----------|----------|
| FolderPath | HasPivFile | CheckedDate |

### Sample Data
| FolderPath | HasPivFile | CheckedDate |
|------------|------------|-------------|
| /sites/MySite/Shared Documents/Department/Team/ProjectA | Yes | 2026-04-08 01:00:00 |
| /sites/MySite/Shared Documents/Department/Team/ProjectB | No | 2026-04-08 01:00:00 |
| /sites/MySite/Shared Documents/Archive/OldProjects/2024 | Yes | 2026-04-08 01:00:00 |

## Steps to Create the Excel Template

### 1. Create New Excel Workbook
1. Go to your SharePoint document library
2. Click **New** > **Excel workbook**
3. Name it: `PIV-Search-Results.xlsx`

### 2. Set Up Headers
In Sheet1, enter the following headers:
- Cell A1: `FolderPath`
- Cell B1: `HasPivFile`
- Cell C1: `CheckedDate`

### 3. Create a Table
1. Select cells A1:C2 (headers + one empty row)
2. Press **Ctrl+T** (or go to Insert > Table)
3. Check "My table has headers"
4. Click **OK**

### 4. Name the Table
1. Click anywhere in the table
2. Go to **Table Design** tab
3. In the "Table Name" field, enter: `PIVResults`
4. Press **Enter**

### 5. Format Columns
- **Column A (FolderPath)**: Text format, widen column
- **Column B (HasPivFile)**: Text format
- **Column C (CheckedDate)**: Date/Time format (right-click > Format Cells > Date)

### 6. Save to SharePoint
Save the file in your SharePoint document library root or a specific folder.

## Power Automate Connection

When connecting Power Automate to this Excel file:

1. Use **"Excel Online (Business)"** connector
2. Action: **"Add a row into a table"**
3. Select:
   - **Location**: Your SharePoint site
   - **Document Library**: Where the Excel file is stored
   - **File**: Browse to `PIV-Search-Results.xlsx`
   - **Table**: Select `PIVResults`

## Optional: Add More Columns

You can extend the table with additional columns:

| Column D | Column E | Column F |
|----------|----------|----------|
| PivFileName | FileCount | Notes |

- **PivFileName**: Name of the PIV file found (if any)
- **FileCount**: Total number of files in the folder
- **Notes**: Any additional comments

If you add columns, update the table range and refresh the table in Power Automate.

## Clearing Old Data (Optional)

If you want to clear old results before each run, add this action before adding new rows:

1. Use **"Delete a row"** action in a loop to remove all existing rows, OR
2. Use **"Send an HTTP request to SharePoint"** to clear the table, OR
3. Create a new Excel file each time with a timestamp in the filename

## Alternative: Use SharePoint List Instead of Excel

If Excel becomes too slow, consider using a SharePoint Custom List:

1. Create a new Custom List in SharePoint
2. Add columns:
   - **FolderPath** (Single line of text)
   - **HasPivFile** (Choice: Yes/No)
   - **CheckedDate** (Date and Time)
3. In Power Automate, use **"Create item"** action for SharePoint Lists

This is often faster and more reliable than Excel for large datasets.
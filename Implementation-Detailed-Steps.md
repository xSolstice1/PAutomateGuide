# Detailed Implementation: Finding Deepest Folders and Checking for PIV Files

## Complete Step-by-Step Power Automate Configuration

This guide provides exact configurations for each action in your Power Automate flow.

---

## Step 1: Trigger
**Action**: "Manually trigger a flow" (Instant cloud flow)

No configuration needed - just click "New step" to continue.

---

## Step 2: Initialize Variables

### Variable 1: Excel File Path
- **Action**: "Initialize variable"
- **Name**: `excelFilePath`
- **Type**: `String`
- **Value**: `/Shared Documents/PIV-Results.xlsx`

### Variable 2: Table Name
- **Action**: "Initialize variable"
- **Name**: `tableName`
- **Type**: `String`
- **Value**: `PIVResults`

> **Note**: No variables needed for the SharePoint site or library — you will select these directly from the dropdown pickers in each SharePoint action.

---

## Step 3: Get All Items from Root Folder

**Action**: "Get files (properties only)" (SharePoint connector)

| Field | Value |
|-------|-------|
| Site Address | Select your SharePoint site from the dropdown |
| Library Name | Select your document library from the dropdown |
| Folder | Browse and select the target folder using the folder picker |
| Limit scope to items inside this folder only | **Uncheck** (we want recursive) |

**Important**: Leave "Limit scope" **unchecked** to get ALL items recursively from the entire library.

---

## Step 4: Filter for Folders Only

**Action**: "Get array items" or use "Filter array"

### Option A: Using Filter Array (Recommended)
**Action**: "Filter array"

| Field | Value |
|-------|-------|
| From | `value` (from "Get files" action) |
| Condition | `ContentType` **contains** `Folder` |

**Alternative Condition**:
- Left side: `IsFolder` (from dynamic content)
- Operator: `is equal to`
- Right side: `true`

This will give you an array of ONLY folders.

---

## Step 5: Process Each Folder

**Action**: "Apply to each"

| Field | Value |
|-------|-------|
| Select an output from previous steps | `Body` from "Filter array" action |

---

## Step 6: Check if Folder is "Deepest" (Has No Subfolders)

Inside the "Apply to each" loop:

**Action**: "Get files (properties only)"

| Field | Value |
|-------|-------|
| Site Address | Select the same SharePoint site from the dropdown |
| Library Name | Select the same document library from the dropdown |
| Folder | `@{items('Apply_to_each')?['FullPath']}` (or use `ServerRelativeUrl`) |
| Limit scope to items inside this folder only | **Check this box** |

This gets all items INSIDE the current folder only.

---

## Step 7: Determine if This is a "Deepest" Folder

**Action**: "Condition"

**Condition Expression**:
```
length(body('Get_files_(properties_only)')?['value']) is equal to 0
```

Or using the count function:
```
length(body('Get_files_(properties_only)')?['value']) equals 0
```

**Logic**: If the folder has NO items inside it, it's a "deepest" (leaf) folder.

**Wait** - This checks if the folder is EMPTY, not if it has no SUBFOLDERS.

### Correct Approach: Check for Subfolders Only

**Action**: "Filter array" (inside the Apply to each)

| Field | Value |
|-------|-------|
| From | `value` from the "Get files" action in Step 6 |
| Condition | `IsFolder` **is equal to** `true` |

Now check if this filtered array is empty:

**Action**: "Condition"

**Condition**:
```
length(body('Filter_array')?['value']) is equal to 0
```

**If Yes**: This folder has NO subfolders = it's a "deepest" folder → Check for PIV files
**If No**: This folder has subfolders → Skip it (or continue processing)

---

## Step 8: Check for PIV Files (Inside "If Yes" Branch)

**Action**: "Get files (properties only)"

| Field | Value |
|-------|-------|
| Site Address | Select the same SharePoint site from the dropdown |
| Library Name | Select the same document library from the dropdown |
| Folder | `@{items('Apply_to_each')?['FullPath']}` |
| Limit scope to items inside this folder only | **Check this box** |
| Filter Query | `substringof('piv', Name) eq true` |

**Important**: The Filter Query will only return files that have "piv" in their name.

---

## Step 9: Determine if PIV File Exists

**Action**: "Condition"

**Condition**:
```
length(body('Get_files_(properties_only)')?['value']) is greater than 0
```

**If Yes**: PIV file found → Set `hasPivFile` to "Yes"
**If No**: No PIV file → Set `hasPivFile` to "No"

---

## Step 10: Add Row to Excel

**Action**: "Add a row into a table" (Excel Online Business connector)

| Field | Value |
|-------|-------|
| Location | Select your SharePoint site |
| Document Library | Select your document library |
| File | `@{variables('excelFilePath')}` |
| Table | `@{variables('tableName')}` |

**Table Fields**:
- **FolderPath**: `@{items('Apply_to_each')?['FullPath']}` (or `ServerRelativeUrl`)
- **HasPivFile**: `Yes` (in "If yes" branch) or `No` (in "If no" branch)
- **CheckedDate**: `@{utcNow()}` or `@{formatDateTime(utcNow(), 'yyyy-MM-dd HH:mm:ss')}`

---

## Complete Flow Structure (Visual)

```
1. Trigger: Manually trigger a flow
   ↓
2. Initialize Variables (excelFilePath, tableName)
   ↓
3. Get files (properties only) - Root folder, NO limit scope (recursive)
   ↓
4. Filter array - Keep only folders (IsFolder equals true)
   ↓
5. Apply to each - Process each folder
   ↓
   6. Get files (properties only) - Current folder, WITH limit scope
      ↓
   7. Filter array - Check for subfolders (IsFolder equals true)
      ↓
   8. Condition - length(filteredArray) equals 0?
      ├─ YES (Deepest folder)
      │   ↓
      │   9. Get files (properties only) - Filter: substringof('piv', Name) eq true
      │   ↓
      │   10. Condition - length(files) greater than 0?
      │       ├─ YES → Add row to Excel: HasPivFile = "Yes"
      │       └─ NO → Add row to Excel: HasPivFile = "No"
      │
      └─ NO (Has subfolders - not deepest)
          └─ Do nothing (skip this folder)
```

---

## Important Notes

### 1. Dynamic Content Names
The exact names in dynamic content may vary. Look for:
- `FullPath` or `ServerRelativeUrl` for folder paths
- `IsFolder` or `ContentType` for folder detection
- `Name` for file names

### 2. Action Names in Expressions
When referencing actions in expressions, use the exact action names:
- `Get_files_(properties_only)` - Note the underscores
- `Filter_array` - Note the underscore
- `Apply_to_each` - Note the underscores

### 3. Handling Large Libraries
If you have thousands of folders:
- Consider adding a delay between iterations
- Use pagination in "Get files" actions
- Consider running during off-peak hours

### 4. Error Handling
Add "Configure run after" settings:
- If any action fails, you can send an email notification
- Add a final "Scope" action with "has failed" condition

### 5. Testing
1. First, test with a small, known folder structure
2. Verify the Excel file is being populated correctly
3. Check the flow run history for any errors

---

## Alternative: Process ALL Folders (Not Just Deepest)

If you want to check EVERY folder (not just deepest), simplify the flow:

```
1. Trigger
   ↓
2. Initialize Variables
   ↓
3. Get files (properties only) - Root, NO limit scope
   ↓
4. Filter array - Keep only folders
   ↓
5. Apply to each - Each folder
   ↓
   6. Get files (properties only) - Current folder, WITH limit scope
      Filter Query: substringof('piv', Name) eq true
      ↓
   7. Condition - length(files) greater than 0?
      ├─ YES → Add row: HasPivFile = "Yes"
      └─ NO → Add row: HasPivFile = "No"
```

This is simpler and may be what you actually need.

---

## Sample Expressions

### Get Folder Path
```
@{items('Apply_to_each')?['FullPath']}
```

### Get Current Date/Time
```
@{formatDateTime(utcNow(), 'yyyy-MM-dd HH:mm:ss')}
```

### Check Array Length
```
@{length(body('Filter_array')?['value'])}
```

### Conditional Value
```
@if(greater(length(body('Get_files_(properties_only)')?['value']), 0), 'Yes', 'No')
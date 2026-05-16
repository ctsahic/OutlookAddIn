# Complete Summary: Fetching Bug Item Metadata from Azure DevOps

## What Was Added

### New Components

**1. AzureDevOpsMetadataService** (`Services/AzureDevOpsMetadataService.cs`)
- Fetches all available fields for Bug work items
- Retrieves work item type information
- Gets field metadata (name, type, usage)
- Prints work item fields to debug output

**2. FieldMetadataDialog** (`Forms/FieldMetadataDialog.cs`)
- Interactive UI dialog for viewing fields
- Displays fields in a sortable data grid
- Shows field name, reference name, type, usage, description
- "Copy to Clipboard" button for easy reference
- Progress indication while loading

**3. "View Fields" Button**
- Added to the ribbon's Settings group
- Located next to "Configuration" button
- Opens the FieldMetadataDialog

### New Classes

```csharp
// Service class
public class AzureDevOpsMetadataService
{
    public async Task<List<WorkItemField>> GetBugFieldsAsync(string pat);
    public async Task<WorkItemType> GetBugWorkItemTypeAsync(string pat);
    public async Task<List<FieldMetadata>> GetBugFieldMetadataAsync(string pat);
    public async Task<WorkItem> GetWorkItemAsync(int workItemId, string pat);
    public void PrintWorkItemFields(WorkItem workItem);
    public string GetWorkItemFieldsAsString(WorkItem workItem);
}

// Data class
public class FieldMetadata
{
    public string ReferenceName { get; set; }
    public string Name { get; set; }
    public string Type { get; set; }
    public string Usage { get; set; }
    public string Description { get; set; }
}
```

## How to Use

### View All Bug Fields in Outlook

1. **Open Outlook**
2. **Click Azure tab** in the ribbon
3. **Click "View Fields"** button
4. **Wait for fields to load** (progress bar shows loading)
5. **Browse all fields** in the grid:
   - Sort by clicking column headers
   - See Field Name, Reference Name, Type, Usage, Description
6. **Copy field names** by:
   - Selecting rows
   - Clicking "Copy Selected to Clipboard"
   - Pasting into your configuration

### Print Fields to Debug Output (Visual Studio)

When creating a work item:
1. All fields automatically print to Debug Output
2. Open Visual Studio ? View ? Output Window
3. Select "Debug" from dropdown
4. Run the add-in and create a bug
5. See complete field list and values

### Access Programmatically

```csharp
var config = new AzureDevOpsConfig { ... };
var metadataService = new AzureDevOpsMetadataService(config);

// Get all fields
var fields = await metadataService.GetBugFieldMetadataAsync(pat);
foreach (var field in fields)
{
    Debug.WriteLine($"{field.Name}: {field.ReferenceName}");
}

// Get work item fields
var workItem = await metadataService.GetWorkItemAsync(123, pat);
metadataService.PrintWorkItemFields(workItem);
```

## Field Information Available

Each field shows:

| Property | Example |
|---|---|
| **Field Name** | Activity |
| **Reference Name** | Microsoft.VSTS.Common.Activity |
| **Type** | String |
| **Usage** | System |
| **Description** | Describes the type of activity for the work item |

## Example Output (Debug Window)

```
========== WORK ITEM FIELDS ==========
Work Item ID: 74
Revision: 1

========== FIELD VALUES ==========
Microsoft.TeamFoundation.Common.InternalFlags: 
Microsoft.VSTS.Common.Activity: A
Microsoft.VSTS.Common.ActivityGroup: tikitak
Microsoft.VSTS.Common.BacklogPriority: 1
Microsoft.VSTS.Common.Priority: 2
Microsoft.VSTS.Common.Severity: 3 - Medium
Microsoft.VSTS.TCM.ReproSteps: <p>test email content</p>
System.AreaPath: Proj1
System.AssignedTo: ...
System.CreatedBy: ...
System.CreatedDate: ...
System.Description: test email content
System.IterationPath: Proj1
System.State: To do
System.Tags: Created-From-Outlook
System.Title: test is not correct
System.WorkItemType: Bug
========== END FIELDS ==========
```

## Use Cases

### 1. Finding Available Custom Fields
**Problem**: You want to know what custom fields your project has
**Solution**: Click "View Fields" and look for fields with Reference Name starting with "Custom."

### 2. Using Field Names in Configuration
**Problem**: You need the exact field name to add to Dynamic Parameters
**Solution**: Click "View Fields", find the field, copy the short name (e.g., "Activity"), add to configuration

### 3. Understanding Field Types
**Problem**: You want to know what type of values a field accepts
**Solution**: Click "View Fields" and check the "Type" column (String, Integer, TreePath, etc.)

### 4. Debugging Field Issues
**Problem**: A field value isn't being set correctly
**Solution**: Create a bug, check Debug Output to see all actual field values and field names

## Prerequisites

- Organization URL configured
- Project Name configured
- Valid PAT token configured
- Internet access to Azure DevOps

## Error Handling

- **"Configuration Required"**: Set Organization URL and Project Name first
- **"PAT Required"**: Enter your Personal Access Token first
- **Network errors**: Check internet connection and Azure DevOps availability
- **Access denied**: Check PAT token permissions and project access

## Files Modified

- `AzureRibbon.cs` - Added viewFields_Click handler
- `AzureRibbon.Designer.cs` - Added viewFieldsButton

## Files Created

- `Services/AzureDevOpsMetadataService.cs` - Metadata fetching service
- `Forms/FieldMetadataDialog.cs` - UI for viewing fields
- `VIEW_FIELDS_GUIDE.md` - User guide
- `FIELD_REFERENCE.md` - Field reference documentation

## Next Steps

1. **Try it out**: Click "View Fields" to see all available fields
2. **Copy field names**: Use the copy button to get reference names
3. **Update configuration**: Add new dynamic parameters based on available fields
4. **Test**: Create bugs with different field combinations
5. **Share with team**: Everyone can now discover available fields!

## Benefits

? **No more guessing** about field names
? **See all available fields** in one place
? **Copy field names** easily for configuration
? **Understand field types** and usage
? **Debug field issues** with print output
? **Discover custom fields** in your project
? **No Azure DevOps** portal access needed

## Technical Details

- Uses Azure DevOps REST API (WorkItemTrackingHttpClient)
- Fetches metadata asynchronously to avoid UI blocking
- Supports all field types: String, Integer, TreePath, Identity, DateTime, Double
- Works with both system and custom fields
- Handles errors gracefully with user-friendly messages

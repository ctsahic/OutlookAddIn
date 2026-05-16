# Quick Reference: Available Azure DevOps Bug Fields

## How to Print All Fields

### Option 1: In Outlook Add-in
1. Open Outlook
2. Click **Azure** tab
3. Click **View Fields** button
4. All Bug fields will be displayed in a sortable table

### Option 2: In Debug Output (Visual Studio)
When you create a work item, all fields are printed to Visual Studio Debug Output:
- Open Visual Studio
- View ? Output Window
- Select "Debug" from dropdown
- Run the add-in and create a bug
- All field names and values appear

### Option 3: Programmatically via API
The `AzureDevOpsMetadataService` class provides these methods:

```csharp
// Get all Bug fields metadata
var metadataService = new AzureDevOpsMetadataService(config);
var fields = await metadataService.GetBugFieldMetadataAsync(pat);

foreach (var field in fields)
{
    Console.WriteLine($"{field.Name} ({field.ReferenceName})");
}

// Get a specific work item's fields
var workItem = await metadataService.GetWorkItemAsync(workItemId, pat);
metadataService.PrintWorkItemFields(workItem);

// Get formatted string
string fieldString = metadataService.GetWorkItemFieldsAsString(workItem);
Debug.WriteLine(fieldString);
```

## Field Structure

Each field has:
- **Name**: Human-readable name (e.g., "Activity")
- **Reference Name**: Full system path (e.g., "Microsoft.VSTS.Common.Activity")
- **Type**: Data type (String, Integer, TreePath, Identity, etc.)
- **Usage**: Whether it's System, Custom, Internal, etc.
- **Description**: What the field is for

## Using Fields in Configuration

### Standard Fields (Supported)
These can be used directly in Dynamic Parameters:

```
Activity ? Microsoft.VSTS.Common.Activity
ActivityGroup ? Microsoft.VSTS.Common.ActivityGroup
Area ? System.AreaPath
Iteration ? System.IterationPath
Priority ? Microsoft.VSTS.Common.Priority
Severity ? Microsoft.VSTS.Common.Severity
State ? System.State
Tags ? System.Tags
Reason ? System.Reason
```

### Custom Fields
Check the "Reference Name" field to see the full path:
- If it starts with "Custom.", use the field name in configuration
- Example: `Custom.MyField` ? Use `MyField` as the key

## Example: Setting Activity Field

1. Click "View Fields"
2. Find row with Name = "Activity"
3. Reference Name shows: `Microsoft.VSTS.Common.Activity`
4. Go to Configuration ? Dynamic Parameters
5. Add: Key = `Activity`, Value = `Development`
6. Create bug ? Activity will be automatically set to "Development"

## Troubleshooting Field Issues

### Field Not Found Error
- Field might not exist in your project's process
- Check with Azure DevOps admin
- Verify correct spelling and case

### Value Not Accepted
- Field value might not be in the allowed values list
- For Activity: Use values like "Development", "Testing", "Deployment"
- Check Azure DevOps for valid values

### Field is Read-Only
- Some system fields can't be set via API
- Try a different field instead
- Contact Azure DevOps admin if needed

## Field Types Explained

- **String**: Text values (Activity, Tags)
- **Integer**: Numbers (Priority - 1, 2, 3, etc.)
- **TreePath**: Hierarchical paths (Area Path, Iteration Path)
- **Identity**: User/Team references (Assigned To)
- **DateTime**: Date and time fields
- **Double**: Decimal numbers

## Common Mistakes

? **Don't:**
- Use the full Reference Name as configuration key
- Use field names with spaces without escaping
- Set read-only system fields
- Use values not in the field's allowed list

? **Do:**
- Use the short name (Activity, Priority, etc.)
- Reference the field metadata dialog to find exact names
- Test one field first before adding multiple
- Check Azure DevOps for valid field values

## Getting Help

1. **View available fields**: Click "View Fields" button
2. **Check reference names**: Look in the Reference Name column
3. **Copy field names**: Right-click and copy from the dialog
4. **Check debug output**: Look at Visual Studio Debug window
5. **Contact support**: Provide the field reference name and error message

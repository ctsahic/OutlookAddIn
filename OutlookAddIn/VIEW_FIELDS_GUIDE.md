# How to View All Bug Item Fields from Azure DevOps

## Overview
You can now view all available fields for Bug work items directly from your Outlook Add-in. This helps you:
- See all field names and reference names
- Understand field types
- Discover custom fields in your Azure DevOps project
- Copy field names for use in configuration

## How to Access Field Metadata

### Method 1: View Available Bug Fields (Recommended)

1. **Open Outlook**
2. **Click the Azure tab** in the ribbon
3. **Click "View Fields"** button in the Settings group
4. A dialog will open showing all available Bug fields

The dialog displays:
- **Field Name**: User-friendly name (e.g., "Activity")
- **Reference Name**: System name used in code (e.g., "Microsoft.VSTS.Common.Activity")
- **Type**: Field data type (String, DateTime, Integer, etc.)
- **Usage**: Field usage context (System, TreePath, Custom, etc.)
- **Description**: Additional field information

### Method 2: View a Specific Work Item's Fields

1. Create a bug using the "Create Bug" button
2. Note the work item ID from the success dialog
3. Go back to the configuration and open "View Fields" dialog
4. The fields from that work item will be displayed in debug output

## Using Field Information

### Copy Reference Names
1. **Select** one or more fields in the list
2. **Click "Copy Selected to Clipboard"**
3. **Paste** into your configuration or code

This helps you understand the exact field names to use.

### Field Types

Common field types you'll see:
- **String**: Text fields (Activity, Tags, etc.)
- **DateTime**: Date fields
- **Double**: Numeric fields
- **Integer**: Whole number fields
- **TreePath**: Hierarchical fields (Area, Iteration)
- **Identity**: User fields (AssignedTo, etc.)

### Using Custom Fields

If your Azure DevOps project has custom fields:
1. They will appear in the list with names like: `Custom.YourFieldName`
2. The Reference Name shows the exact path
3. You can use these in your dynamic parameters configuration

## Example Workflow

**Scenario**: You want to see all available fields and understand which ones you can use in configuration.

1. **Setup Prerequisites**:
   - Enter Organization URL (e.g., `https://dev.azure.com/yourorg`)
   - Enter Project Name
   - Enter PAT token

2. **View Fields**:
   - Click "Azure" tab
   - Click "View Fields"
   - See all available fields

3. **Identify Usable Fields**:
   - Look for fields with **Usage** = "System" or "Custom"
   - Avoid "Internal" or "Computed" fields
   - Standard fields like Activity, Priority, Severity are good choices

4. **Copy Field Names**:
   - Select fields you want to use
   - Click "Copy Selected to Clipboard"
   - Use the Reference Name in your configuration

5. **Configure**:
   - Go to "Configuration" button
   - Add Dynamic Parameters using the field names
   - Example: Key = `Activity`, Value = `Development`

## Debug Output

When you open a work item's field view, detailed field information is printed to:
- **Visual Studio Debug Output Window**
- **Or**: View ? Output window ? Debug option

This includes:
- Work Item ID
- All field reference names
- All field values
- Complete field structure

## Supported Fields (Common Examples)

| Field Name | Reference Name | Type | Usage |
|---|---|---|---|
| Activity | Microsoft.VSTS.Common.Activity | String | System |
| Activity Group | Microsoft.VSTS.Common.ActivityGroup | String | System |
| Assigned To | System.AssignedTo | Identity | System |
| Area Path | System.AreaPath | TreePath | System |
| Iteration Path | System.IterationPath | TreePath | System |
| Priority | Microsoft.VSTS.Common.Priority | Integer | System |
| Severity | Microsoft.VSTS.Common.Severity | String | System |
| State | System.State | String | System |
| Tags | System.Tags | String | System |
| Title | System.Title | String | System |
| Description | System.Description | String | System |

## Troubleshooting

### "Configuration Required" Error
- Make sure you've set Organization URL and Project Name
- Enter them in the ribbon fields
- Click "Create Bug" once to save them

### "PAT Required" Error
- Make sure you've entered your Personal Access Token in the PAT field
- Click "Create Bug" once to save it

### No Fields Appear
- Check your internet connection
- Verify PAT token is still valid
- Verify Organization URL is correct
- Verify Project Name exists and you have access

### Can't Find a Specific Field
- Custom fields might use different names
- Check with your Azure DevOps administrator
- Look in Azure DevOps project settings ? Process for custom fields

## Next Steps

After viewing fields:
1. **Identify fields** you want to populate automatically
2. **Copy the Reference Name** for each field
3. **Use simple names** in your configuration (last part of the path)
4. **Test with one field** before adding multiple
5. **Add to Dynamic Parameters** in Configuration dialog

Example:
- Reference Name: `Microsoft.VSTS.Common.Activity`
- Use in Config: Key = `Activity`, Value = `Development`

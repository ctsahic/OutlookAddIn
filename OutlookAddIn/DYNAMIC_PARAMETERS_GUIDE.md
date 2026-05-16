# Dynamic Parameters and Custom Fields Guide

## How Dynamic Parameters Are Now Used

When you create a bug, all dynamic parameters stored in the Configuration dialog are automatically applied to the work item as field values.

### Supported Field Mappings

The system automatically maps parameter keys to Azure DevOps field paths:

#### **Standard Azure DevOps Fields**
| Parameter Key | Maps To | Azure DevOps Field |
|---|---|---|
| `Activity` | `/fields/Microsoft.VSTS.Common.Activity` | Activity |
| `ActivityGroup` | `/fields/Microsoft.VSTS.Common.ActivityGroup` | Activity Group |
| `Area` | `/fields/System.AreaPath` | Area Path |
| `AreaCode` | `/fields/System.AreaPath` | Area Path |
| `Iteration` | `/fields/System.IterationPath` | Iteration Path |
| `Priority` | `/fields/Microsoft.VSTS.Common.Priority` | Priority |
| `Severity` | `/fields/Microsoft.VSTS.Common.Severity` | Severity |
| `State` | `/fields/System.State` | State |
| `Tags` | `/fields/System.Tags` | Tags |
| `Reason` | `/fields/System.Reason` | Reason |

#### **Custom Fields**
For any field name that's not in the standard mappings, the system will try:
1. First check if it's a known field (from the table above)
2. If not found, construct a path as `/fields/Custom.{FieldName}`

### Examples

#### **Example 1: Set Activity**
1. Open Configuration ? Dynamic Parameters
2. Add Parameter:
   - Key: `Activity`
   - Value: `bbb`
3. Click Save
4. When you create a bug, the Activity field will be set to `bbb`

#### **Example 2: Set Multiple Fields**
1. Open Configuration ? Dynamic Parameters
2. Add multiple parameters:
   - Key: `Activity`, Value: `Development`
   - Key: `Priority`, Value: `High`
   - Key: `Severity`, Value: `Critical`
3. Click Save
4. When creating a bug, all three fields will be populated automatically

#### **Example 3: Custom Field**
If you have a custom field in your Azure DevOps project:
1. Open Configuration ? Dynamic Parameters
2. Add Parameter:
   - Key: `CustomFieldName`
   - Value: `CustomValue`
3. Click Save
4. The system will create the field path as `/fields/Custom.CustomFieldName`

### How Dot-Notation Works

If you use dot notation like `ActivityGroup.Activity`:
- The system extracts the last part: `Activity`
- Maps it to: `/fields/Microsoft.VSTS.Common.Activity`
- This allows flexible naming while still mapping to the correct field

### Workflow Example

**Scenario**: You want to automatically set Activity to "bbb" every time you create a bug.

1. **Setup (One-time)**
   - Click "Configuration" button
   - Go to "Dynamic Parameters" tab
   - Click "Add Parameter"
   - Enter Key: `Activity`
   - Enter Value: `bbb`
   - Click "Save"

2. **Create Bug (Every time)**
   - Select an email in Outlook
   - Click "Create Bug" button
   - The bug will be created with Activity automatically set to `bbb`
   - No need to manually enter it!

### Error Handling

If a field doesn't exist in your Azure DevOps project:
- The bug will still be created (bypassRules is enabled)
- Fields that don't exist will be skipped silently
- The required fields (Title, Description) will always be populated

### Best Practices

? **Do:**
- Use simple, recognizable field names
- Keep values appropriate for the field type
- Test with one parameter first before adding many

? **Don't:**
- Use special characters in field names
- Set contradictory values (e.g., Severity and Priority conflicts)
- Leave values empty if the field is mandatory

### Technical Details

The parameter mapping happens in the `MapParameterKeyToFieldPath()` method which:
1. Takes the parameter key
2. Extracts the field name (last part if dot-notation)
3. Looks up the field in the standard mappings
4. Falls back to custom field path if not found
5. Returns the Azure DevOps field path for the JSON Patch Document

All dynamic parameters are sent to Azure DevOps in a single PATCH request for efficiency.

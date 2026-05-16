# Azure DevOps Configuration System - User Guide

## How to Add Key-Value Parameters

### Step 1: Open Configuration Dialog
1. Open Outlook
2. Look at the **Ribbon** (top menu bar)
3. Click on the **"Azure"** tab
4. In the **"Settings"** group, click **"Configuration"** button

### Step 2: Add Dynamic Parameters
1. The **Configuration Dialog** will open
2. Go to the **"Dynamic Parameters"** tab
3. Click **"Add Parameter"** button
4. A new row will be added to the grid
5. Enter your key-value pair:
   - **Key**: The parameter name (e.g., "custom_field", "area_code")
   - **Value**: The parameter value (e.g., "Bug", "Team1")

### Step 3: How Dynamic Parameters Are Used
The dynamic parameters can be used to:

1. **Populate Missing Fields**
   - If a mandatory field is empty and a value exists in the key-value configuration with the matching key, it will be populated automatically

2. **With "Ignore Invalid" Flag**
   - Fields marked with "Ignore Invalid = true" will use dynamic parameters even if validation errors occur
   - This allows fields to be optional but still populated if configured

### Step 4: Configure Field Behavior
1. Go to the **"Field Configuration"** tab
2. View/Edit field settings:
   - **Field Name**: The name of the field
   - **Display Name**: User-friendly name
   - **Configuration Key**: The key that maps to a dynamic parameter
   - **Mandatory**: Whether the field is required
   - **Ignore Invalid**: If true, allows field to be optional but populated from dynamic parameters
   - **Default Value**: Default value if nothing else is provided

### Example Workflow

**Scenario**: You want to always set the Area to "Team1" without typing it every time.

1. Open Configuration ? Dynamic Parameters tab
2. Click "Add Parameter"
3. Enter:
   - Key: `area_code`
   - Value: `Team1`
4. Click "Save"

Now, when creating a bug:
- If the Area field is empty, it will automatically be populated with "Team1" from the key-value configuration
- Users can still override it by typing a different value

## Key Features

? **Dynamic Mandatory Parameters**: Define which fields are required through configuration  
? **Ignore Invalid Support**: Fields can be marked to ignore validation and still populate from key-value config  
? **Secure Storage**: All configuration stored in Windows Credential Manager  
? **Fallback Chain**: Current Value ? Dynamic Parameters ? Default Value  
? **Easy to Update**: No need to modify code - just add/edit parameters in the Configuration dialog

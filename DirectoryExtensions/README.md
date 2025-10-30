# Microsoft Entra Directory Extension Management Functions

This repository contains PowerShell helper functions to manage **custom directory extensions (schema extensions)** in Microsoft Entra ID (Azure AD).
The functions allow you to **create, read, update, and remove** directory extensions for applications and users, with proper permission checks and tenant-wide considerations.

---

## Helper Functions Overview

### 1. `Test-Module`

**Purpose:**  
Ensures that a required PowerShell module is installed and imported before running dependent functions.

**Features:**
- Checks if the module is already loaded.  
- Tries to import it automatically if not.  
- Installs it from PowerShell Gallery if missing.  
- Provides verbose, friendly console output for debugging.  
- Handles long import times for large modules like `Microsoft.Graph` and `Az`.

**Usage Example:**
```powershell
Test-Module -Name Microsoft.Graph.Authentication
```

### 2. `igall` — Microsoft Graph Pagination Helper

**Purpose:**

The `igall` function is a PowerShell helper designed to simplify retrieving **all paginated results** from Microsoft Graph API endpoints using `Invoke-MgGraphRequest`.  
It automatically follows `@odata.nextLink` pages and supports the `ConsistencyLevel: eventual` header for advanced queries.

**Usage Examples:**

Get all users from Microsoft Graph
```Powershell
igall "https://graph.microsoft.com/v1.0/users"
```

Get all users with advanced consistency
```powershell
igall "https://graph.microsoft.com/beta/users?$count=true" -Eventual
```

Limit to firt 5 pages 

```powershell 
igall "https://graph.microsoft.com/v1.0/users" -limit 5
```



### 3. `Show-AvailableFunctions` — List All Functions in the Script
#### Purpose
The `Show-AvailableFunctions` command lists all functions included in the helper script and provides a brief description for each.
It serves as a quick reference to discover available tools and understand their purpose without opening the script file.

When the script is loaded, a short message is displayed reminding users that this command is available.

#### Behavior
- Automatically lists all PowerShell functions defined in the script.

- Optionally expands each function to show its inline help comment block (if available).

- Displays function names in **green** and descriptions in **gray** for clarity.

- The helper message automatically appears after importing or dot-sourcing the script.

#### Required Permissions
None — this function runs locally and doesn’t require Microsoft Graph access or any elevated privileges.

#### Usage Exmpale

```powershell
# List all functions in the helper script
Show-AvailableFunctions
```
##### Output Exmaple
```powershell
📜 Available Directory Extension Functions:

• Get-DirectoryExtensions
  ↳ Lists all directory extensions on one or all registered Entra ID applications.

• Get-DirectoryExtensionValues
  ↳ Fetches directory extension values for one or all users.

• New-DirectoryExtensionForUser
  ↳ Creates a custom directory extension (schema extension) in Entra ID for user objects.

• Remove-ApplicationDirectoryExtension
  ↳ Safely removes one or more extension properties from Microsoft Entra applications.

• Set-DirectoryExtensionValue
  ↳ Sets a specific directory extension value for a user in Microsoft Entra ID.

• Show-AvailableFunctions
  ↳ Lists all custom functions available in the current script and shows their brief help.

Tip: Use 'Get-Help <FunctionName> -Full' to see detailed documentation.
```
## Functions Overview
### 1. `New-DirectoryExtensionForUser` — Create a Directory Extension for Users in Entra ID

####  Purpose
The `New-DirectoryExtensionForUser` function creates a **custom directory extension (schema extension)** on a registered Microsoft Entra (Azure AD) application.  
These extensions can store additional user metadata — for example, HR codes, employee types, or project-specific attributes — that aren't part of the default user schema.

####  Required Permissions
- **Delegated Microsoft Graph scopes:** `Application.ReadWrite.All`

If your session does not already include it, the function automatically reconnects using that scope.
- **Entra ID Roles:**  
  - Application Administrator  
  - Cloud Application Administrator  
  - Global Administrator  

#### Parameters
| Name | Type | Required | Description |
|------|------|-----------|-------------|
| `ApplicationObjectID` | String | ✅ | The Object ID of the registered application to attach the directory extension to. |
| `NameOfExtension` | String | ✅ | The name of your directory extension (e.g. `UserPurpose`). |

#### Usage Example
```powershell
New-DirectoryExtensionForUser `
  -ApplicationObjectID "11111111-2222-3333-4444-555555555555" `
  -NameOfExtension "UserPurpose"
  ```

#### Output

✅ Connected to Microsoft Graph with scope 'Application.ReadWrite.All'.
Let's create a new directory extension called 'UserPurpose' on application ID 6d3e8f4a-2d42-40b0-8e8a-2b90b2b558f4...
✅ Successfully created extension: extension_b3137f118a934bd288018d5c873ebaf7_UserPurpose


### 2. `Get-DirectoryExtensions` — List Directory Extensions in Entra ID

#### Purpose
The `Get-DirectoryExtensions` function retrieves **directory extensions (schema extensions)** registered on one or more **Microsoft Entra ID (Azure AD)** application objects.  
These extensions define custom attributes that can be assigned to users, groups, or other directory objects.

---

#### Required Permissions
**Microsoft Graph delegated scopes:**
- `Application.Read.All`

If your session doesn’t already include it, the function will automatically reconnect using the correct scope.

**Entra ID roles (any of the following):**
- Application Administrator  
- Cloud Application Administrator  
- Global Reader  

---

#### Parameters

| Name                 | Type     | Description                                                                                                                                      | Required |
| -------------------- | -------- | ------------------------------------------------------------------------------------------------------------------------------------------------ | -------- |
| **`AppDisplayName`** | `String` | The display name of the application whose directory extensions should be listed. If omitted, the function scans **all registered applications**. | ❌        |


---

#### Examples

```powershell
# Retrieve directory extensions for a specific app
Get-DirectoryExtensions -AppDisplayName "Custom Identity App"

# Retrieve all directory extensions from every registered app
Get-DirectoryExtensions

```

#### Output

| Property            | Description                                                          |
| ------------------- | -------------------------------------------------------------------- |
| **ApplicationName** | The display name of the application.                                 |
| **ApplicationID**   | The Object ID of the application.                                    |
| **ExtensionName**   | The full directory extension name.                                   |
| **ExtensionID**     | The unique identifier for the extension.                             |
| **TargetObjects**   | The object types (e.g., `User`, `Group`) that can use the extension. |

### 3. `Get-DirectoryExtensionValues` — Retrieve Directory Extension Values from Users in Entra ID

Fetches directory extension values for one or all users in Microsoft Entra ID.

---

#### Purpose

Purpose

The Get-DirectoryExtensionValues function retrieves values from custom directory extensions (schema extensions) applied to Entra ID user objects.
It can return:

- All users with any extension value present, or

- Values for a specific extension, or

- Values for a single user.

This function is ideal for auditing, reporting, or validating custom Entra ID attributes (e.g. HRCode, CostCenter, or UserPurpose).

---

#### Required Permissions

- **Delegated Microsoft Graph scopes:** User.Read.All

If the current session does not include the required permission, the function automatically reconnects using the proper scope.

- **Entra ID Roles:**

  - User Administrator


#### Parameters

| Name                     | Type   | Required | Description                                                                                                                                                                                   |
| ------------------------ | ------ | -------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `DirectoryExtensionName` | String | ❌        | The name of a specific directory extension to retrieve (e.g. `extension_b3137f118a934bd288018d5c873ebaf7_UserPurpose`). If not specified, all directory extensions with values are retrieved. |
| `UserUPN`                | String | ❌        | The UPN of a single user to target. If omitted, the function loops through all users in the tenant.                                                                                           |

#### Usage Examples

Retrieve all directory extension values for all users:

```powershell
Get-DirectoryExtensionValues
``` 
Retrieve a specific directory extension for all users:
```Powershell
Get-DirectoryExtensionValues -DirectoryExtensionName "extension_b3137f118a934bd288018d5c873ebaf7_UserPurpose"
```
Retrieve all directory extension values for one user:
```Powershell
Get-DirectoryExtensionValues -UserUPN "jane.doe@domain.com"
```
Retrieve a specific directory extension for one user:
```Powershell
Get-DirectoryExtensionValues `
  -DirectoryExtensionName "extension_b3137f118a934bd288018d5c873ebaf7_UserPurpose" `
  -UserUPN "john.smith@domain.com"
```

#### Output
```Powershell
✅ Connected to Microsoft Graph with scope 'User.Read.All'.
📋 Retrieving directory extension values for all users...
✅ Finished collecting extension values.
```
Example result:
| DisplayName | UserPrincipalName                                     | ExtensionName                                          | ExtensionValue |
| ----------- | ----------------------------------------------------- | ------------------------------------------------------ | -------------- |
| Jane Doe    | [jane.doe@domain.com](mailto:jane.doe@domain.com)     | extension_b3137f118a934bd288018d5c873ebaf7_HRCode      | DK1001         |
| John Smith  | [john.smith@domain.com](mailto:john.smith@domain.com) | extension_b3137f118a934bd288018d5c873ebaf7_UserPurpose | MobileUser     |


### 4. `Set-DirectoryExtensionValue` — Update Directory Extension Values for Users in Entra ID

#### Purpose
The `Set-DirectoryExtensionValue` function updates the value of a custom directory extension (schema extension) for a user object in Microsoft Entra ID.
It’s useful when you need to maintain or modify user-specific metadata such as DepartmentCode, Responsibility, or HRStatus — attributes that are not part of the default schema.

#### Required Permissions
- **Delegated Microsoft Graph scopes:** `Directory.ReadWrite.All`

If your current Graph session doesn’t already include this scope, the function will automatically reconnect using the proper permission.

- **Entra ID Roles:** 
  - User Administrator

  #### Parameters

| Name                     | Type   | Required | Description                                                                                               |
| ------------------------ | ------ | -------- | --------------------------------------------------------------------------------------------------------- |
| `DirectoryExtensionName` | String | ✅        | The full name of the directory extension (e.g. `extension_b3137f118a934bd288018d5c873ebaf7_UserPurpose`). |
| `UserUPN`                | String | ✅        | The User Principal Name (UPN) of the target user whose extension value will be updated.                   |
| `NewValue`               | String | ✅        | The new value to assign to the specified directory extension.                                             |

#### Usage Examples

Set a directory extension value for a user:

```powershell
Set-DirectoryExtensionValue `
  -DirectoryExtensionName "extension_b3137f118a934bd288018d5c873ebaf7_UserPurpose" `
  -UserUPN "alex.jensen@domain.com" `
  -NewValue "MobileUser"
```
Update another custom attribute:

```Powershell
Set-DirectoryExtensionValue `
  -DirectoryExtensionName "extension_1234567890abcdef_HRCode" `
  -UserUPN "mia.larsen@domain.com" `
  -NewValue "DK1001"
```

#### Output

```Powershell
✅ Connected to Microsoft Graph with scope 'Directory.ReadWrite.All'.
✏️ Setting 'extension_b3137f118a934bd288018d5c873ebaf7_UserPurpose' for user 'alex.jensen@domain.com
' to 'MobileUser'...
✅ Successfully updated extension value for alex.jensen@domain.com
.
```
If an error occurs:

```Powershell
❌ Failed to update extension value for alex.jensen@domain.com.
Error: Insufficient privileges to complete the operation.
```
### 5. `Remove-ApplicationDirectoryExtension` — Safely Remove Directory Extensions from Applications in Entra ID

#### Purpose
The `Remove-ApplicationDirectoryExtension` function safely removes one or more **custom directory extensions** (schema extensions) from Microsoft Entra (Azure AD) applications.
It ensures that the extension exists, prompts for administrator confirmation, and warns that the removal will affect **all users in the tenant** that rely on the extension.

This is particularly useful when cleaning up unused extensions or decommissioning application attributes.

#### Required Permissions
- **Delegated Microsoft Graph scopes:** `Application.ReadWrite.All`

If your current Graph session doesn’t already include this scope, the function will automatically reconnect using the proper permission.

- **Entra ID Roles:** 
  - Application Administrator
  - Cloud Application Administrator
  - Global Administrator

#### Parameters

| Name            | Type     | Required | Description                                                                                          |
| --------------- | -------- | -------- | ---------------------------------------------------------------------------------------------------- |
| `ApplicationId` | String[] | ✅        | One or more Object IDs or App IDs of the applications to remove the extension from. Can be piped in. |
| `ExtensionId`   | String   | ✅        | The ID of the extension property to remove.                                                          |

#### Usage Examples
Remove a directory extension from a single application:

```Powershell
Remove-ApplicationDirectoryExtension `
  -ApplicationId "11111111-2222-3333-4444-555555555555" `
  -ExtensionId "abcd1234-5678-90ef-ghij-1234567890kl"
```


#### Output

```Powershell
✅ Connected to Microsoft Graph with scope 'Application.ReadWrite.All'.
Checking for extension property on application: 11111111-2222-3333-4444-555555555555
✅ Extension property found:
Name: extension_b3137f118a934bd288018d5c873ebaf7_UserPurpose
ID: abcd1234-5678-90ef-ghij-1234567890kl
Target App: 11111111-2222-3333-4444-555555555555

⚠️ This will remove the extension property for ALL users in the tenant that rely on it. Are you sure you want to continue? (Y/N)

✅ Extension property successfully removed from application 11111111-2222-3333-4444-555555555555.

```
If the extension is not found:

```Powershell
❌ No extension property found with ID 'abcd1234-5678-90ef-ghij-1234567890kl' on application 11111111-2222-3333-4444-555555555555.

```



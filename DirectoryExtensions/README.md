# Directory Extension Utility for Entra ID

This PowerShell module provides a set of helper functions to **verify module availability**, **connect to Microsoft Graph**, and **create or inspect custom directory extensions** in Entra ID (formerly Azure AD).

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
## Functions Overview
### 1. `New-DirectoryExtensionForUser`

####  Purpose
Creates a **custom directory extension** (schema extension) on an Entra ID (Azure AD) application object.  
This is useful for storing additional user metadata that isn’t part of the default schema — for example, `UserPurpose`, `EmployeeType`, or `HRCode`.

####  Required Permissions
- **Delegated Microsoft Graph scopes:** `Application.ReadWrite.All`
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

### 2. `Get-DirectoryExtensions`

#### Purpose
Retrieves all **directory (schema) extensions** registered on one or more Entra ID (Azure AD) applications.  
If an application display name is provided, only that app’s extensions are returned.  
If omitted, it enumerates all applications in the tenant and lists their extensions.

---

#### Required Permissions
**Microsoft Graph delegated scopes:**
- `Application.Read.All`

**Entra ID roles (any of the following):**
- Application Administrator  
- Cloud Application Administrator  
- Global Reader  

---

#### Parameters

| Name | Type | Required | Description |
|------|------|-----------|-------------|
| `AppDisplayName` | String | ❌ | The display name of the Entra ID application. If omitted, retrieves extensions from all applications. |

---

#### Examples

```powershell
# Retrieve directory extensions for a specific app
Get-DirectoryExtensions -AppDisplayName "Custom Identity App"

# Retrieve all directory extensions from every registered app
Get-DirectoryExtensions

```
### 3. `Get-DirectoryExtensionValues`

Fetches the values of a specific **directory extension** (custom schema attribute) for one or all Entra ID users.

---

#### Description

This function queries **Microsoft Graph** to read user attributes that were created as **directory extensions**  
(e.g., `extension_<AppId>_HRCode`, `extension_<AppId>_EmployeeType`).

You can use it to:
- Retrieve the value of an extension for a **specific user**.
- Retrieve all users’ values for auditing or reporting.

---

#### Syntax

```powershell
Get-DirectoryExtensionValues -DirectoryExtensionName <String> [-UserUPN <String>]
```


# Azure AD & Entra Role Reporting Script

##  Overview

This PowerShell script collects and exports comprehensive reports of **Entra ID** and **Azure role assignments** into an Excel file.  
It retrieves both **eligible** (PIM-based) and **active** administrator assignments across Microsoft Entra ID and Azure subscriptions.

The resulting report includes:
- Directory administrators and their assigned roles
- Privileged Identity Management (PIM) eligible role assignments
- Azure subscription role assignments (via `Az` module)

---

##  Features

- Automatically installs and imports required PowerShell modules:
  - `Az.Accounts`
  - `Az.Resources`
  - `Microsoft.Graph.Authentication`
  - `ImportExcel`
- Handles **paged Graph API responses** using the custom `igall` helper function
- Fetches detailed user attributes (sign-in info, creation date, etc.)
- Expands **group-based role assignments** and tracks progress with `Write-Progress`
- Filters out **temporary PIM activations** to avoid duplicates in the active admin list
- Exports data to **Excel** with multiple worksheets:
  - `Administrators`
  - `Eligible Roles`
  - `Azure Roles`
- Provides clear console feedback and progress indicators during execution

---

##  Requirements

- PowerShell 7 or later
- The following PowerShell modules:
  - `Az.Accounts`
  - `Az.Resources`
  - `Microsoft.Graph.Authentication`
  - `ImportExcel`
- Sufficient Microsoft Entra ID and Azure permissions:
  - `RoleManagement.Read.Directory`
  - `User.Read.All`
  - `Directory.Read.All`
  - `GroupMember.Read.All`
  - `RoleEligibilitySchedule.Read.Directory`
  - `RoleManagement.Read.All`
  - `AuditLog.Read.All`
- Reader roles to both Entra ID and Azure Subscriptions:
  - `Global Reader`
  - `Az Reader Role`

> ⚠️ **Note:** The script uses the Microsoft Graph **beta endpoint**, which may be subject to change.

---

##  Usage

1. **Run PowerShell**

2. **Execute the script**

   ```powershell
   .\Generate-EntraAdminReport.ps1

3. **Authentication** 
- You will be prompted to sigin in twice:
    - Once for your Azure Account (Connect-AzAcccount)
    - Once for your Microsoft Graph (Connect-MgGraph)
4. Select export folder
-The script will display:
```sql
⚠️  The folder selection window may appear behind other open windows.
```
If you don’t see it immediately, minimize other windows.
5. Wait for the script to complete
- Progress and details will be shown in the console (with colors and progress bars).
6. Output 
- The resulting Excel report will be saved as:
```php-template
<OrgName>-EntraIDAdminReport<YYYY-MM-DD>.xlsx
```

## Output Structure

The script generates an Excel workbook named:
```php-template
<OrgName>-EntraIDAdminReport<YYYY-MM-DD>.xlsx
```

### Workbook Overview

| Worksheet | Description |
|------------|--------------|
| **Administrators** | Lists all active users who are currently assigned to Azure AD directory roles. This excludes users who only have *eligible* roles via PIM (Privileged Identity Management). |
| **Eligible Roles** | Contains all users and groups that are *eligible* for elevated roles through PIM. Group-based eligibilities are automatically expanded to include all members. |
| **Azure Roles** | Lists all users, service principals, and groups with assigned Azure subscription-level roles (RBAC). Each record includes the role, scope (subscription), and object type. |

###  Columns

#### **Administrators**
Contains all currently active directory role members.

| Column | Description |
|--------|-------------|
| `Role` | Name of the assigned directory role. |
| `DisplayName` | Full name of the user. |
| `UserPrincipalName` | User’s sign-in name (UPN). |
| `CompanyName` | Company or organization name (from Entra ID). |
| `AccountEnabled` | Indicates whether the account is active. |
| `CreatedDateTime` | Date the user account was created. |
| `LastPasswordChangeDateTime` | Last time the user changed their password. |
| `LastSignInDateTime` | Most recent successful sign-in (if available). |

---

#### **Eligible Roles**
Lists all users and groups *eligible* for directory roles via PIM (Privileged Identity Management).

| Column | Description |
|--------|-------------|
| `DisplayName` | User’s full name. |
| `UserPrincipalName` | User’s sign-in name (UPN). |
| `EligibleRole` | Name of the eligible role. |
| `DirectRole` | Indicates if the role is directly assigned. |
| `EligibleRoleGroup` | Group name from which the role is inherited (if applicable). |
| `MemberType` | Type of membership (e.g., Direct, Group). |
| `CreatedDateTime` | Date the user account was created. |
| `LastPasswordChangeDateTime` | Last time the user changed their password. |
| `LastSignInDateTime` | Most recent successful sign-in (if available). |

---

#### **Azure Roles**
Contains Azure subscription-level RBAC assignments.

| Column | Description |
|--------|-------------|
| `RoleDefinitionName` | Name of the Azure RBAC role. |
| `DisplayName` | Name of the assigned user, group, or service principal. |
| `SigninName` | Sign-in name (for users). |
| `ObjectId` | Object ID of the assigned principal. |
| `Subscription` | Name of the Azure subscription where the role applies. |


##  Functions Overview

The script defines several helper functions to handle pagination, caching, and module management for Microsoft Graph and Azure operations.

---

###  `igall`
Fetches all paginated results from Microsoft Graph API calls.

**Parameters:**
- `Uri` — The full Graph API endpoint URL.

**Description:**
Microsoft Graph often returns results in pages.  
`igall` automatically follows the `@odata.nextLink` property to retrieve all pages and return a complete dataset.

**Example:**
```powershell
$users = igall -Uri "https://graph.microsoft.com/beta/users"
```

### `Get-User`

Retrieves detailed user information from Microsoft Graph with built-in caching to minimize redundant API calls.

**Parameters:**
- `Id` — *(String)* The Azure AD Object ID of the user.

**Description:**
The function first checks whether the requested user is already stored in the `$cache` hashtable.  
If not found, it queries the Microsoft Graph **beta endpoint** for the user and selects key properties including:

- `DisplayName`
- `UserPrincipalName`
- `companyName`
- `accountEnabled`
- `CreatedDateTime`
- `LastPasswordChangeDateTime`
- `signInActivity`
- `lastSignInDateTime`
- `lastNonInteractiveSignInDateTime`

This ensures efficient, low-latency lookups when multiple references to the same user occur throughout the script.

**Example:**
```powershell
$user = Get-User -Id "a1b2c3d4-5678-90ef-ghij-klmnopqrstuv"
Write-Host "Fetched user: $($user.DisplayName)"
```
### `Test-Module`

Ensures that a PowerShell module is available by verifying, importing, and installing it if necessary.

**Parameters:**
- `Name` — *(String)* The name of the module to validate, import, or install.

---

**Description:**
The `Test-Module` function is a self-healing helper designed to make sure all required modules are ready before execution.  
It performs the following steps:

1. **Checks if the module is already imported.**
2. **Attempts to import** the module if it’s installed but not loaded.
3. **Automatically installs** the module from the PowerShell Gallery if it’s missing.
4. Provides **clear, color-coded output** to inform the user of each step.
5. Includes **special handling** for modules like:
   - `Microsoft.Graph` (import time notice)
   - `Az` (long import time warning)

---

**Example Usage:**
```powershell
Test-Module -Name "Microsoft.Graph.Authentication"
Test-Module -Name "Az.Accounts"
Test-Module -Name "ImportExcel"

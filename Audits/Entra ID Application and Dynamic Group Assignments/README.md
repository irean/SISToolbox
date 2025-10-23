# Entra ID Application and Dynamic Group Assignment Exporter

This PowerShell script is designed to provide a comprehensive audit of Entra ID (Azure AD) user access and memberships. Its primary purpose is to help identity and access management teams, auditors, and IT administrators understand which users have access to which applications and which dynamic groups they belong to.  

The script performs the following functions:

- Disconnects any existing Microsoft Graph sessions to ensure a clean connection.
- Checks that all required modules (`Microsoft.Graph.Authentication` and `ImportExcel`) are installed and imported, installing them automatically if missing.
- Connects to Microsoft Graph with the necessary delegated permissions to read users, groups, applications, and assignments.
- Allows targeting specific users via an Excel file or exporting all users of type `member`.
- Retrieves application role assignments for each user, including detection of SSO configuration and SCIM synchronization.
- Retrieves dynamic group memberships and membership rules for users.
- Exports all results to a structured Excel file with separate worksheets for application assignments and dynamic group memberships.

**Purpose of the Script:**

- To provide visibility into which users have access to which applications and groups.  
- To enable auditing and governance of Entra ID environments.  
- To support security reviews, compliance reporting, and access management initiatives.  
- To document user access and group membership for ongoing IT operations or migration planning.  

This script is especially useful for organizations that need to ensure proper access controls, verify dynamic group memberships, or analyze the configuration of application assignments in their Entra ID tenant.
## Functions

### Function: igall
Retrieves all paginated results from Microsoft Graph.

**Parameters**

| Parameter   | Description |
|------------|-------------|
| `-Uri`     | The Microsoft Graph API endpoint |
| `-Eventual`| Adds the `ConsistencyLevel: eventual` header |
| `-limit`   | Limits pagination depth (default: 1000) |

**Example Usage**

```powershell
$results = igall -Uri "https://graph.microsoft.com/v1.0/users"
```

### Function: ConvertTo-PSCustomObject

Recursively converts nested hashtables into PowerShell objects for easier handling of Graph API responses.

**Parameters**

| Parameter   | Description |
|------------|-------------|
| `-InputObject`     | A hashtable or array of hashtables from Graph API response |

**Example Usage**

```Powershell
$psObject = ConvertTo-PSCustomObject -InputObject $hash
```

**Function: test-module**

Ensures a PowerShell module is installed and imported. If missing, it installs automatically.

| Parameter   | Description |
|------------|-------------|
| `-Name`     | The name of the PowerShell module to check/import |

**Example Usage**

```powershell
Test-Module -Name "Microsoft.Graph.Authentication"
```

## Required Permissions

The script requires the following Microsoft Graph delegated permissions:`

```pgsql
Organization.Read.All
User.Read.All
GroupMember.Read.All
Group.Read.All
Application.Read.All

```
## Dependencies

* Microsoft.Graph.Authentication
* ImportExcel

The script automatically installs missing modules for the current user.

## Usage Instructions

1. Open PowerShell (preferably as Administrator).

2. Run the script:

```powershell
.\Export-EntraID_AppRole_DynamicGroups.ps1
```
3. Sign in to Microsoft Graph when prompted.

4. Specify whether to use an Excel file of users:
   - Type `Y` to select a file containing a `userPrincipalName` column.
   - Type `N` to target all users of type `member`.

5. If using an Excel file, select the file using the file picker dialog.

6. Choose a folder where the results will be saved using the folder picker dialog.

7. Wait for the script to complete. The script will export an Excel file containing two worksheets:
   - `ApplicationAssignments` – contains user application role assignments including SSO and SCIM information.
   - `DynamicGroupAssignments` – contains dynamic group memberships and membership rules.

8. Open the exported Excel file to review the audit results.

### Output
The Excel file will be saved as: 
```php-template
<OrganizationName>_Application_Dynamicgroup_Assignments_<YYYY-MM-DD>.xlsx
```


# Entra ID Access Package Assignment Audit
A PowerShell toolkit for auditing and comparing **Microsoft Entra ID (Azure AD)** users against **Access Package assignments** in **Entra ID Governance**.  

This script automates the process of:
- Connecting to Microsoft Graph  
- Retrieving users dynamically (via filters, group membership, or manual lists)  
- Comparing users against a specific Access Package  
- Displaying real-time progress and exporting results  

## Features

- Automatically ensures required PowerShell modules (`Microsoft.Graph.*`, `ImportExcel`) are installed and imported  
- Retrieve Entra ID users by:
  - Country, Department, Company Name, EmployeeId prefix  
  - Group membership  
  - Manual CSV import  
- Compare user lists with Access Package assignments in Microsoft Graph    
- Supports Entra ID Governance use cases  

## Requirements

| Component | Minimum Version | Notes |
|------------|----------------|-------|
| PowerShell | 7.2+ | Recommended for cross-platform compatibility |
| Modules | Microsoft.Graph, Microsoft.Graph.Authentication, ImportExcel | Installed automatically if missing |
| Permissions | `User.Read.All`, `EntitlementManagement.Read.All`, `Group.Read.All`, `Organization.Read.All` | Required for Graph API access |

---

## Installation

1. Clone or download this script.
2. Open PowerShell and run as Administrator.
3. Import the script:

   ```powershell
   . .\Start-UserAccessPackageAudit.ps1
   ```
4. Run the main command:
    ```powershell
    Start-UserAccessPackageAudit
    ```
## Workflow Overview
1. **Module Validation** 

Ensures all required modules are present using Test-Module.

2. **Graph Connection**

Authenticates via Connect-MgGraph with the required scopes.

3. **User Retrieval Options**

Choose one of:

    - Filter by country, department, company, etc.

    - Retrieve group members

    - Import users manually via CSV

4. **Access Package Comparison**

Compares each user against Access Package assignments with detailed progress and output.

5. **Export Results**

Selects a folder for report export via Windows folder picker.

## Function Overview

1. `Test-Module`
Ensures that a required PowerShell module is **installed** and **imported**.  
If the module is missing, it will automatically install it from the PowerShell Gallery and import it into the current session.
2. `ConvertTo-PSCustomObject`
Recursively converts nested hashtables or arrays into structured PSCustomObject instances.

Useful when working with complex JSON from Graph API responses.
3. `igall`
Handles Microsoft Graph pagination automatically.
```powershell
igall -Uri "https://graph.microsoft.com/v1.0/users" -Eventual
```
- Adds ConsistencyLevel=eventual when specified.

- Retrieves all pages of data with a loop up to a limit.
4. `Get-AccessPackageAssignmentsTargets`
Retrieves all **Delivered** Access Package assignments for a given **Access Package ID**.
```powershell
Get-AccessPackageAssignmentsTargets -AccessPackageId "b3a77f84-6a3d-44b1-9f50-d32c17346a31"
```
Outputs user details such as:
- AccessPackageName

- Policy Name

- Target Object ID

- PrincipalName

- Assignment State and Status
5. `Compare-UsersToAccessPackageAssignmentsWithProgress`
Compares a list of users against Access Package assignments.
```powershell
Compare-UsersToAccessPackageAssignmentsWithProgress -AccessPackageId "b3a77f84-6a3d-44b1-9f50-d32c17346a31" -UserList $users
```
6. `Get-UsersDynamic`

Flexible user retrieval for different scenarios.

Syntax: 
```powershell
Get-UsersDynamic [[-Country] <String>] [[-Department] <String>] [[-CompanyName] <String>] [[-GroupId] <String>] [[-EmployeeIdStartsWith] <String>] [-EmployeeIdNotNull] [-EmployeeLeaveDateTimeNotNull]
```
7. `Select-FolderPath`

Opens a Windows dialog for selecting a folder where reports will be saved.
8. `Start-UserAccessPackageAudit`
The main function that ties everything together.
```powershell
Start-UserAccessPackageAudit
```

Prompts interactively for:

- Connection to Graph

- User retrieval method

- Access Package ID

- Export folder

Then runs the comparison and generates a report.

## Example Usage    

```powershell 
PS C:\code> start-UserAccessPackageAudit
==========================================
🔎  USER & ACCESS PACKAGE AUDIT STARTED
==========================================
Checking module 'Microsoft.Graph.Authentication'...
✅ Module 'Microsoft.Graph.Authentication' is already imported.
Checking module 'ImportExcel'...
✅ Module 'ImportExcel' is already imported.

 Connecting to Microsoft Graph...
✅ Connected to Graph successfully.

Fetching organization name...
🏢 Organization: Company Name 

How would you like to get users?
1️⃣  Filter (Country, EmployeeId, etc.)
2️⃣  Group (Members of specific group)
3️⃣  Manual (CSV import)
Enter choice (1-3): 1

--- 🔍 FILTER PARAMETERS ---
Enter country (or leave blank):
Filter by department (or leave blank): 
Filter by companyName (or leave blank): Company Name
Filter EmployeeId not null? (y/n): y
Filter EmployeeLeaveDateTime not null? (y/n): n
🔎 Querying users with: companyName eq 'Company Name' and accountEnabled eq true

✅ Retrieved 13 users for comparison.

Enter Access Package ObjectId: 58f12289-ca67-46bd-aebe-222b6ff3877e
📦 Selected Access Package: Workvivo

🔍 Comparing users to Access Package assignments...

 Starting user comparison against Access Package assignments...
Querying Microsoft Graph for assignments...                                                                             
    → https://graph.microsoft.com/v1.0/identityGovernance/entitlementManagement/assignments?$expand=target,accessPackage,assignmentPolicy&$filter=accessPackage/id eq '58f12289-ca67-46bd-aebe-222b6ff3877e' and state eq 'Delivered'
✅ Retrieved 1573 Access Package assignments.
👥 Beginning comparison for 13 users...
✅ Comparison complete for 13 users.
Summary:
   → Assigned: 13
   → Not assigned: 0

📁 Select output folder for Excel export...
--------------------------------------------------------
📂 Please select a folder where the report will be saved.
⚠️  The folder selection window may appear behind other open windows.
If you don't see it, try minimizing other windows.
--------------------------------------------------------
✅ Export folder selected: C:\Reports

DetailedReport   : DetailedReport
Summary          : Summary
Package          : OfficeOpenXml.Packaging.ZipPackage
Encryption       : OfficeOpenXml.ExcelEncryption
Workbook         : OfficeOpenXml.ExcelWorkbook
DoAdjustDrawings : True
File             : C:\Reports folder\Company_Name-Workvivo-2025-11-10.xlsx
Stream           : System.IO.MemoryStream
Compression      : Level6
Compatibility    : OfficeOpenXml.Compatibility.CompatibilitySettings

✅ Excel report created successfully: C:\Reports folder\Company_Name-Workvivo-2025-11-10.xlsx
Sheets included: Summary, DetailedReport, UsersNotAssigned
==========================================
✅  USER & ACCESS PACKAGE AUDIT COMPLETE
==========================================

```
## Required Microsft Graph Scopes

The following **delegated permissions** are required:

- `User.Read.All`  
- `Directory.Read.All`  
- `EntitlementManagement.Read.All`  
- `Group.Read.All`  
- `Organization.Read.All`  

>  If not already consented, you’ll be prompted during the Graph connection.
---

## 🧑‍💻 Author & Metadata

| Field | Info |
|-------|------|
| **Author** | Sandra Saluti |
| **Version** | 1.0 |
| **Date** | 2025-11-10 |
| **Tags** | Microsoft Graph, Entra ID, Governance, Access Packages |
| **License** | MIT |


---
## References

- [Microsoft Graph API – Entitlement Management](https://learn.microsoft.com/en-us/graph/api/resources/entitlementmanagement-overview)  
- [Microsoft Graph API – Access Package Assignment](https://learn.microsoft.com/en-us/graph/api/resources/accesspackageassignment?view=graph-rest-1.0)  
- [Microsoft Graph PowerShell SDK](https://learn.microsoft.com/en-us/powershell/microsoftgraph/overview)

---
## Example CSV Format for Manual Import

When using **manual import mode**, the CSV should be formatted like this:

```csv
ObjectId,DisplayName,PrincipalName,EmployeeId,EmployeeHireDate,JobTitle,AccountEnabled,Department,CompanyName
8c5ab4e3-14a2-47f8-a6a3-9ef8c97d3a90,Jane Doe,jane.doe@company.com,12345,2024-01-01,Consultant,True,IT,Company Name
``` 

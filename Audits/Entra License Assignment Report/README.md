# License Reporting Script Documentation

## Overview

This PowerShell script is designed to generate reports on license usage within Microsoft Graph for specified tenants. It connects to the Microsoft Graph API to gather information about assigned licenses, users, and devices, and then generates detailed reports in Excel format. The script aims to provide insights into license allocation and utilization across different organizations.

## Prerequisites

- PowerShell version 7 or later
- Granted permissions for script accessing Microsoft Graph Api
- Microsoft Graph PowerShell module (`Microsoft.Graph.Authentication`)
- Microsoft Graph PowerShell module (`Microsoft.Graph.Users`)
- Microsoft Graph PowerShell module (`Microsoft.Graph.DeviceManagement`)
- ImportExcel module

## Usage

### Importing Required Modules

The script ensures that necessary PowerShell modules are installed and imported by function Test-Module. This script relies on two external modules: `licenses.ps1` and `companies.ps1`. Ensure that they are downloaded in same directory.

### Connecting to tenants

The script automatically connects to tenants listed in companies.ps1 with function `Connect-LicenseReport`

#### Parameter: Organization
This parameter validates companynames in companies.ps1 but is not required. If omitted, the script will process all the companies listed.

Example: Single Tenant Report
```powershell
Connect-LicenseReport -Organization "your-company-name"
```
### Get-LicenseReport
Generates  license usage for specified  tenants.

Example: All Tenants in list

```powershell
Get-LicenseReport
```

### OutPut

The script generates Excel reports containing the following information:

- License allocation and utilization breakdown.
- User details including display name, email, employee ID, job title, city, department, company, and more.
- Device details including ID, name, owner type, operating system, last sync date, compliance state, and OS version.

Tabs included:

1. UserData - Contains data about the users
      - UserPrincipalName
      -  EmployeeID
      -  EmployeeNumber
     -   SKU
      -  License (friendly name)
      -  DisplayName
      -  Title
      -  City
      -  StreetAddress
      -  Country
      -  Company
      -  Department
      -  Usertype
      -  Enabled
      -  Created
      -  LastSignInDatetime
      -  LastNonInteractiveSiginDateTime
  1. License - Contains list of users with assigned licenses
  2. LicensePivotTable -  a pivot with current licenses in the tenant
  3. LicenseActivated - Contains a table that captures mismatch in the amount of licenses bought for tenant and licenses assigned to users
   4. Mailboxusers - Contains list users with a mailbox enalbed license, used to define how many accounts that are/need backup
   5. Devices - List of Intune managed devices in tenant
   6. Device Count - Pivot counting the devices based of operatingSystem


#### Output save path
The script checks if sharepoint directory is synchronized and uploads there
1. If you want the report to automatically upload to Teams, follow the steps below. Otherwise, skip this and the file will be saved in `Documents\Reports`.
   1. Visit folder/ Sharepoint and create a folder that matches you companyname defined in the companies.ps1
   2. If you want to upload automatically to a specified teams folder, navigate there and hit the sync. The file will automatically save new files to this directory if matched in settings
```powershell
$env:USERPROFILE\<savepath>\'orgname'"
```

If path not available script saves to 

```PowerShell
$env:USERPROFILE\Documents\Reports\
```



## Updating StandardLicense Friendlynames


This report retrieves the **SkupartNumber** from the Entra ID and associates it with a **Friendly name** sourced from the file Licenses.ps1. If the **Skupartnumber** doesn't match any friendly name in the file, the result will display **"unknown"** under the **License** column in the excel report.

### Adding a New License Friendly Name

1. Visit [Microsoft's Licensing Service Plan Reference](https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference) and locate the **SkupartNumber** in the **"string id"** column along with the corresponding **Friendly name** under **"product name"**.

2. Open the file **`licenses.ps1`** and insert a new line in the script just above the line stating **"Unknown"**.


```powershell


        "" {  "Unlicensed" }

```
3. Enter the new **product name** and the **skupartnumber** to be matched.

```powershell


"New_SKuPartNumber" {  "New Friendly Name" }

        "" {  "Unlicensed" }
```

4.  Save the file.


Following these steps ensures that all license reports will now include the new skupartnumber and match it with a friendly product name.

## Adding a New Customer

When adding a new customer report, follow these steps:

1. Open `companies.ps1` and copy existing user line
    Example:
    ```powershell
    'Customername' = @{TenantID = 'b87c6fe7-b729-466c-b6fc-21ad1cf5d9e1'; clientID = '2b42314b-21db-4247-9d59-4bb08a60c609'} ```



2. Update the 'CustomerName' with the new customers name. This does not need to be exactly as the tenant org name. The name will generate filename and the path if autosave to sharepoint or your documents. 

3. Update number in TenantID with the new customers tenantID in Entra ID.

4. Create a new appregistration in tenant. Name it cmcms-report.Copy  the clientID and add in string after clientID .

5. Add scopes as application
- User.read.all
- Directory.read.all
- Policy.read.all
- Auditlog.read.all
- Mailboxsettings.read
- DeviceManagementConfiguration.Read.All
- DeviceManagementManagedDevices.PrivilegedOperations.All
- DeviceManagementManagedDevices.Read.All
- Group.Read.All
- GroupMember.Read.All
- Organization.Read.All

6. Create a client secret and copy the value
7. Depending on automation, save these values in keyvault, passwordsaver or what you might prefer
   - TenantID
   - ClientID
   - clientSecret
8. Test connection by running script.














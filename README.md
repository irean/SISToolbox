# SISToolbox
Collection of scripts that I find useful in my daily work

****

##  SetManagedIdentityPermissions

### Overview

The `SetManagedIdentityPermissions` function is a powerful tool for managing application permissions granted to a managed identity. This documentation provides an overview of the function, including its included sub-functions, and usage examples.

### Included Functions

#### 1. Test-Module

The `Test-Module` function is included to simplify the process of connecting and using `Set-MIPermissions`. It verifies the presence of required modules and ensures that your environment is properly configured for managing managed identity permissions.

#### 2. Set-MIPermissions

The `Set-MIPermissions` function is the core of this toolkit. It connects to Microsoft Graph with specific scopes, allowing you to manage application permissions assigned to a managed identity effectively.

### Usage

#### Script Initialization

Before using the `SetManagedIdentityPermissions` functions, make sure you have the necessary modules installed. These functions require access to Microsoft Graph with the following scopes:

- `Application.Read.All`
- `ApproleAssignment.Readwrite.All`
- `RoleManagement.Readwrite.Directory`

#### Parameters

- **ManagedIdentityID**: This parameter specifies the unique identifier (ID) of the managed identity to which you want to grant or modify application permissions.

- **Roles**: Specify the roles or permissions to assign to the managed identity. You can assign multiple roles by separating them with commas.

### Example

Here's an example of how to use the `Set-MIPermissions` function:

```powershell
Set-MIPermissions -ManagedIdentityId <id> -Roles 'user.read.all', 'directory.read.all'
``` 



## Function: TestModuleImport

### Overview

The `Test-Module` function is a useful utility script designed to verify the presence of required modules and ensure their proper importation. This script simplifies the process of checking module dependencies before running actual scripts or connecting to APIs, enhancing script reliability and reducing errors related to missing or unimported modules.

## Function Purpose

- **Module Dependency Verification**: `Test-Module` serves as a lightweight script for confirming that essential modules are installed and ready for use.

- **Enhanced Script Reliability**: By incorporating this function into your scripts, you can significantly enhance their reliability and robustness, preventing issues related to missing or unimported modules.

## Usage

To check for the presence and importation of a specific module, simply invoke the `TestModuleImport` function with the `-name` parameter followed by the name of the module you wish to validate. Here is an example:

```powershell
Test-Module -name Microsoft.Graph
```

## Script: SaveFileLocation

### Overview

The `SaveFileLocation` script is a handy utility designed to facilitate the selection of a folder path by opening the Windows File Explorer. This script is particularly useful in situations where you want to provide users with the flexibility to choose where to save a file. It ensures the verification of the selected folder and stores the path in the `$folder` variable for further use in your scripts.

### Script Purpose

- **Folder Selection**: `SaveFileLocation` opens the Windows File Explorer, allowing users to browse and select a folder where they want to save a file.

- **Path Verification**: After the user selects a folder, the script verifies the chosen path to ensure that the users has selected a folder.

- **Path Storage**: The selected folder path is stored in the `$folder` variable, making it available for use in your scripts.


## Script: New-MgSecurityGroupDynamicUser

### Overview

The `New-MgSecurityGroupDynamicUser` script is a tool for creating dynamic user security groups using Microsoft Graph, targeting specific licenses by their service plans. This script captures all users assigned a specific license to a security group. 

## Script Purpose

- **Dynamic User Security Groups**: This script automates the creation of dynamic user security groups based on the selected license.

- **Service Plan Targeting**: You can specify the Service Plan ID (ServicePlanID) to include in the dynamic group. Refer to Microsoft's official documentation for [Product names and service plan identifiers for licensing](https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference) to find the relevant Service Plan IDs.

- **Exclusion Support**: Optionally, you can provide exclusion Service Plan IDs (notincludeServicePlanID) to separate one license from another. This helps ensure precise targeting within the selected license.

- **Organization Context**: The organization name (orgnames) is incorporated into the security group name for naming purpose. The resulting group name follows the format: `lic-audit-companyname-licensename`.

## Parameters

- **$licname**: This parameter specifies the name of the license you want to target, for example, "M365_E3".

- **$ServicePlanID**: Provide the ID of the service plan you want to include in the dynamic user security group. Refer to the Microsoft documentation mentioned above for Service Plan IDs.

- **$notincludeServicePlanID**: (Optional) Use this parameter to specify exclusion IDs that help separate one license from another. It ensures precise targeting within the selected license.

- **$orgnames**: This parameter represents the organization name and is used to create the security group name in the format: `lic-audit-companyname-licensename`.

## Example

```powershell
# Create a dynamic user security group targeting specific licenses and service plans

# Specify the license name, service plan ID, and organization name
$licname = "M365_E3"
$ServicePlanID = "<ServicePlanID>"
$orgnames = "companyname"

# Optionally, specify exclusion Service Plan IDs
$notincludeServicePlanID = "<ExclusionServicePlanID>"

# Invoke the New-MgSecurityGroupDynamicUser script
New-MgSecurityGroupDynamicUser -licname $licname -ServicePlanID $ServicePlanID -notincludeServicePlanID $notincludeServicePlanID -orgnames $orgnames





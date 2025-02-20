

# OneDrive Sharing Permissions Remover

This script is designed to remove sharing permissions on all files and folders in a user's OneDrive for Business (ODFB) environment. It utilizes the Microsoft Graph API and requires certain permissions to interact with users' OneDrive data.

The script is based on Vasil Michevs script that removes access from one user, you can find that script here [Github Post](https://github.com/michevnew/PowerShell/blob/master/Graph_ODFB_remove_all_shared.md)

## Requirements

- PowerShell version 7.0 or higher
- The AppID used must have the following permissions:
  - `User.Read.All` to enumerate all users in the tenant
  - `Sites.ReadWrite.All` to return all the item sharing details

For more details on permissions, refer to [this blog post](https://www.michev.info/blog/post/3018/remove-sharing-permissions-on-all-files-in-users-onedrive-for-business).

## Script Overview

This script removes sharing permissions on files and folders within users' OneDrive for Business. It works by enumerating all users in the tenant and checking each user's OneDrive for shared files or folders, then removing the sharing permissions.

### Parameters

- `$TenantID` *(string)*: The ID of the tenant where the users' OneDrive for Business resides.
- `$ClientSecretCredential` *(pscredential)*: The credential to authenticate against Microsoft Graph.
- `$ExpandFolders` *(switch, default: $true)*: Determines whether to expand folders and include their items in the output.
- `$Depth` *(int, default: 2)*: Specifies the folder depth for expanding or including items in nested folders.

### Functions

#### `igall`
A helper function that makes repeated requests to the Graph API to handle paginated responses. It fetches all items from the given URI.

#### `processChildren`
Processes the child items (files, folders, and notebooks) within a specified folder, checking if they are shared. If shared, permissions are removed.

#### `processFolder`
Processes an individual folder, checking for shared status. If the folder is shared, it removes the permissions and continues to process its child items if necessary.

#### `RemovePermissions`
Removes sharing permissions for the specified item in the user's OneDrive.

### Script Flow

1. Connect to Microsoft Graph using the provided `TenantID` and `ClientSecretCredential`.
2. Iterate through all users in the tenant.
3. For each user, check if OneDrive is provisioned and retrieve their files and folders.
4. Remove sharing permissions on shared files, folders, or notebooks.
5. Output a list of items with removed sharing permissions.

### Example Usage

```powershell
# Define the necessary parameters
$TenantID = "your-tenant-id"
$ClientSecretCredential = Get-Credential

# Run the script
.\Remove-OneDriveSharingPermissions.ps1 -TenantID $TenantID -ClientSecretCredential $ClientSecretCredential -ExpandFolders -Depth 2


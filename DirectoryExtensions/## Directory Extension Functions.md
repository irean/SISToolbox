# Directory Extension Functions

The main goal of these functions is to create data extensions that are discoverable and filterable throughout the tenant. However, there are some limitations, such as a maximum of 100 extensions and difficulties in locating extensions from accidentally deleted applications.

For more in-depth information, please refer to the [Microsoft documentation](https://learn.microsoft.com/en-us/graph/extensibility-overview?tabs=http).

## Define Directory Extension

1. **Create an App Registration**
    - Go to Azure AD > App Registrations > Create a new registration.
    - Add the following permissions as application permissions: "user.readwrite.all" and "directory.read.all".

2. **Create the Directory Extension**
    1. Connect to Microsoft Graph with the "application.readwrite.all" scope.
    2. Visit the Repository User LifeCycles and download the PowerShell Script `directoryextensions.ps1`.
    3. Run the script in PowerShell 7.
    4. Execute the function `Create-DirectoryExtensionUserPurpose` with the following parameters:
       - `-ApplicationObjectID` (the Object ID of the application created in step 1).
       - `-nameofextension` (the desired extension name).

3. **Add Values to the New Extension**
    - You can add values to the new extension from logic apps, scripts, and more.

### Example

```powershell
PS C:\Code\GitHub\reports> Create-DirectoryExtension -ApplicationObjectID -nameofextension UserPurposeExtension
```    

## List Directory Extensions
1. Connect to your tenant using PowerShell 7 with the applications.read.all permission.

2. If not already downloaded, fetch and run the directoryextensions.ps1 script from GitHub.

3. Run the function List-DirectoryExtensions to list all application directory extensions. You can use the optional parameter -AppDisplayname to filter extensions for a specific app.

### Example 
```powershell
PS C:\Code\GitHub\reports> List-DirectoryExtension -AppDisplayname 'Room-attributes'

Name                           Value
----                           -----
appDisplayName
name                           extension_e79b8f8d3e3f4fd39f8008084c154fc8_userPurposeExtension
id                             d9a6fdd7-b137-4304-9a1d-e9c15513a3a9
deletedDateTime
dataType                       String
isSyncedFromOnPremises         False
targetObjects                  {User}

```

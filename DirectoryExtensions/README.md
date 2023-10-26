## Directory Extension Functions

The main purpose is to create data extensions that are disoverable and filterabel through out the tenant. There are however a few limitations like 100 extensions and problem to find extensions from applications that were accidently deleted. 

More information can be found here https://learn.microsoft.com/en-us/graph/extensibility-overview?tabs=http

### Define Directory Extension

1. Create an app registration
    1. Go to Azure AD -> App Registrations -> Create new
    1. Add permissions "user.readwrite.all" , "directory.read.all" as application
1. Create the Directory Extension
   1.  Connect to Mg with scopes application.readwrite.all
    2. Go to Repository User LifeCycles and download PowerShell Script directoryextensions.ps1
    3. Run file  in __Powershell 7__
    4. Run function __Create-DirectoryExtensionUserPurpose -ApplicationObjectID__ _<the objectID of the application you created in step 1>_ __-nameofextension__ _< the name you want to use >_
2. Add value to the new extension from logic apps, scripts and more 

#### Example

```
PS C:\Code\GitHub\reports> Create-DirectoryExtension -ApplicationObjectID -nameofextension UserPurposeExtension

```




### List directory extensions

1. Connect to tenant with applications.read.all in __Powershell 7__
2. If not downloaded before, download and run directoryextensions.ps1 from GitHub
3. Run function List-DirectoryExtensions
    1. __List-DirectoryExtensions__ returns all application directory extensions
    1. __List-DicretoryExtensions -AppDisplayname__ _< displayname of app >_  returns the chosen apps directory extensions.
#### Example 

```
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

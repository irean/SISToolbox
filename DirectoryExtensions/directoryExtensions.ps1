function test-module {
    [CmdletBinding()]
    param(
        [String]$Name
      
    )
    Write-Host "Checking module $name"
    if (-not (Get-Module $Name)) {
        Write-Host "Module $Name not imported, trying to import"
        try {
            if ($Name -eq 'Microsoft.Graph') {
                Write-Host "Microsoft.Graph module import takes a while"
                Import-Module $Name  -ErrorAction Stop
            }
            elseif ($Name -eq 'Az') {
                Write-Host "Module Az is being imported. This might take a while"
            }
            else {
                Import-Module $Name  -ErrorAction Stop
            }
            
        }
        catch {
            Write-Host "Module $Name not found, trying to install"
            Install-Module $Name -Scope CurrentUser -AllowClobber -Force -AcceptLicense -SkipPublisherCheck
            Write-Host "Importing module  $Name "
            Import-Module $Name  -ErrorAction stop 
    
        }
    } 
    else {
        Write-Host "Module $Name is imported"
    }   
}



Test-Module -Name  Microsoft.Graph.Authentication

<# 
.SYNOPSIS
Creates a custom directory extension (schema extension) in Entra ID for user objects.

.DESCRIPTION
This function helps you create a directory extension property on an Entra ID Application object.
Useful when you need to store additional user metadata not covered by default attributes.

.ROLE
Application Admini, Cloud App Admin

.PARAMETER ApplicationObjectID
The Object ID of the registered application where the directory extension will be created.

.PARAMETER NameOfExtension
The name of your directory extension, for example: "UserPurpose" or "HRCode".
#>

function New-DirectoryExtensionForUser {
    param (
        [Parameter()]
        [String]$ApplicationObjectID,
        [Parameter()]
        [String]$nameofextension
    )
Connect-MgGraph -Scopes "Application.ReadWrite.All"
    Write-Host "Let's create a new directory extension called '$NameOfExtension' on application ID $ApplicationObjectID..." -ForegroundColor Cyan

    $body = @{
        name          = $nameofextension
        datatype      = "String"
        targetObjects = @(
            "User"
        )
    }
    try {
        $response = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/applications/$ApplicationObjectID/extensionProperties" -Body $body
        Write-Host "✅ Successfully created extension: $($response.name)" -ForegroundColor Green
    }
    catch {
        Write-Host "❌ Something went wrong while creating the directory extension:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor DarkRed
    }
}

<# 
.SYNOPSIS
Lists all directory extensions on one or all registered Entra ID applications.

.DESCRIPTION
If you provide an application display name, it lists extensions for that app only.
Otherwise, it loops through all registered applications and shows their extensions.

.ROLE
Application Admini, Cloud App Admin

.PARAMETER AppDisplayName
The display name of the target application (optional).
#>
function Get-DirectoryExtensions {
    param (
        [Parameter()]
        [String]$AppDisplayName
    )

    Connect-MGgraph -scopes "Application.Read.All"
    Write-Host "🔍 Checking directory extensions..." -ForegroundColor Cyan

    if ($AppDisplayname) {
        Write-Host "Looking up extensions for application: '$AppDisplayName'..."
        $app = Invoke-MggraphRequest -Method GET "https://graph.microsoft.com/v1.0/applications?`$Filter=displayname eq '$AppDisplayname'" | Select-Object -ExpandProperty Value | Select-OBject -Expandproperty id 
        if ($app) {
            Invoke-MgGraphRequest -Method GET "https://graph.microsoft.com/v1.0/applications/$app/extensionProperties" |
            Select-Object -ExpandProperty value
        }
        else {
            Write-Host "⚠️ No application found with that name."
        }
    }
    else {
        Write-Host "No specific app provided – scanning all registered applications..."
        $url = "https://graph.microsoft.com/v1.0/applications"

        do {
            $response = Invoke-MgGraphRequest -Method GET -Uri $url
            $url = $response.'@odata.nextLink'

            foreach ($app in $response.value) {
                Write-Host "→ Checking extensions for: $($app.displayName)"
                $ext = Invoke-MgGraphRequest -Method GET "https://graph.microsoft.com/v1.0/applications/$($app.id)/extensionProperties"
                $ext.value | Select-Object name, id, targetObjects
            }

        } while ($url)
    }

    Write-Host "✅ Done listing directory extensions." -ForegroundColor Green
}

<# 
.SYNOPSIS
Fetches values of a specific directory extension for one or all users.

.DESCRIPTION
If you specify a UserUPN, it retrieves the extension value for that user.
If not, it loops through all users and displays the value.

.Role
User Administrator

.PARAMETER DirectoryExtensionName
The name of the directory extension (must be the full schema name).

.PARAMETER UserUPN
The UPN of a specific user (optional).
#>


function Get-DirectoryExtensionValues {
    param (
        [Parameter()]
        [String]$DirectoryExtensionName,
        [Parameter()]$UserUPN
    )

    Connect-MgGraph -Scopes "User.Read.All"

    if ($UserUPN) {
        Write-Host "👤 Getting '$DirectoryExtensionName' value for user $UserUPN..." -ForegroundColor Cyan
        $user = Invoke-MgGraphRequest -Method GET "https://graph.microsoft.com/beta/users/$UserUPN?`$select=id,displayName,$DirectoryExtensionName"
        $user | Select-Object id, displayName, $DirectoryExtensionName
    }
    else {
        Write-Host "📋 No user specified — retrieving '$DirectoryExtensionName' for all users (this might take a while)..." -ForegroundColor Yellow
        $url = "https://graph.microsoft.com/beta/users?`$select=id,displayName,userPrincipalName,$DirectoryExtensionName"

        do {
            $response = Invoke-MgGraphRequest -Method GET -Uri $url
            $url = $response.'@odata.nextLink'

            foreach ($user in $response.value) {
                [PSCustomObject]@{
                    ID                = $user.id
                    DisplayName       = $user.displayName
                    UserPrincipalName = $user.userPrincipalName
                    ExtensionValue    = $user.$DirectoryExtensionName
                }
            }
        } while ($url)

        Write-Host "✅ Finished collecting extension values." -ForegroundColor Green
    }
}
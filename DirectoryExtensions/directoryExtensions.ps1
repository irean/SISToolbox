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
function ConvertTo-PSCustomObject {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true, Mandatory = $true)]
        [System.Collections.Hashtable] $InputObject
    )
    Process {
        if ( $InputObject) {
            $o = New-Object psobject
            foreach ($key in $InputObject.Keys) {
                $value = $InputObject[$key]
                if ($value -and $value.GetType().FullName -match 'System.Object\[\]') {
                    if ($value.Count -gt 0 -and $value[0].GetType().FullName -match 'System.Collections.Hashtable') {
                        $tempVal = $value | ConvertTo-PSCustomObject
                        Add-Member -InputObject $o -NotePropertyName $key -NotePropertyValue $tempVal
                    }
                    elseif ($value.Count -gt 0 -and $value[0].GetType().FullName -match 'System.String') {
                        $tempVal = $value | ForEach-Object { $_ }
                        Add-Member -InputObject $o -NotePropertyName $key -NotePropertyValue $tempVal
                    }
                }
                elseif ($value -and $value.GetType().FullName -match 'System.Collections.Hashtable') {
                    Add-Member -InputObject $o -NotePropertyName $key -NotePropertyValue (ConvertTo-PSCustomObject -InputObject $value)
                }
                else {
                    Add-Member -InputObject $o -NotePropertyName $key -NotePropertyValue $value
                }
            }

            Write-Output $o
        }
    }
}

function igall {
    [CmdletBinding()]
    param (
        [string]$Uri,
        [switch]$Eventual,
        [int]$limit = 1000
    )
    $nextUri = $uri
    $count = 0
    $headers = @{
        Accept = 'application/json'
    }
    if ($Eventual) {
        $headers.Add('ConsistencyLevel', 'eventual')
    }
    do {
        $result = Invoke-MgGraphRequest -Method GET -uri $nextUri -Headers $headers
        $nextUri = $result.'@odata.nextLink'
        if ($result.value) {
            $result.value | ConvertTo-PSCustomObject
        }
        elseif ($result.value -and $result.value.GetType().FullName -match 'System.Object\[\]') {
            @()
        }
        elseif ($result) {
            $result | ConvertTo-PSCustomObject
        }
        $count += 1
    } while ($nextUri -and ($count -lt $limit))
}
function Show-AvailableFunctions {
    <#
    .SYNOPSIS
    Lists all custom functions available in the current script and shows their brief help.
    .DESCRIPTION
    Uses `Get-Command` and `Get-Help` to dynamically display each function's name and `.SYNOPSIS` line.
    #>

    Write-Host "`n📜 Available Directory Extension Functions:`n" -ForegroundColor Cyan

    # Collect only your own functions (e.g. containing 'DirectoryExtension' or adjust as needed)
    $functions = Get-Command -CommandType Function |
                 Where-Object { $_.Name -match 'DirectoryExtension' -or $_.Name -match 'Show-AvailableFunctions' }

    if (-not $functions) {
        Write-Host "⚠️ No matching functions found." -ForegroundColor Yellow
        return
    }

    $functions | ForEach-Object {
        $name = $_.Name
        # Extract .SYNOPSIS line if available
        $synopsis = (Get-Help $name -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Synopsis -ErrorAction SilentlyContinue)
        if (-not $synopsis) { $synopsis = "(No help available)" }

        Write-Host "• $name" -ForegroundColor Green
        Write-Host "  ↳ $synopsis" -ForegroundColor DarkGray
        Write-Host ""
    }

    Write-Host "Tip: Use 'Get-Help <FunctionName> -Full' to see detailed documentation." -ForegroundColor Cyan
}



Test-Module -Name  Microsoft.Graph.Authentication

Write-Host ""
Write-Host "📘 Helper loaded successfully!" -ForegroundColor Green
Write-Host "Use 'Show-AvailableFunctions' to list all available functions and their descriptions." -ForegroundColor Cyan
Write-Host ""



function New-DirectoryExtensionForUser {
    param (
        [Parameter()]
        [String]$ApplicationObjectID,
        [Parameter()]
        [String]$nameofextension
    )
    $context = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $context) {
        Write-Host "🔗 Not connected to Microsoft Graph. Connecting with required scope..." -ForegroundColor Cyan
        Connect-MgGraph -Scopes "Application.ReadWrite.All"
    }
    else {
        $scopes = $context.Scopes
        if ('Application.ReadWrite.All' -notin $scopes) {
            Write-Host "⚠️ Current Graph session does not include required scope 'Application.ReadWrite.All'." -ForegroundColor Yellow
            Write-Host "🔁 Reconnecting with proper permissions..." -ForegroundColor Cyan
            Disconnect-MgGraph | Out-Null
            Connect-MgGraph -Scopes "Application.ReadWrite.All"
        }
        else {
            Write-Host "✅ Connected to Microsoft Graph with scope 'Application.ReadWrite.All'." -ForegroundColor Green
        }
    }
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
}


function Get-DirectoryExtensions {


    param (
        [Parameter()]
        [string]$AppDisplayName
    )
    $context = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $context) {
        Write-Host "🔗 Not connected to Microsoft Graph. Connecting with required scope..." -ForegroundColor Cyan
        Connect-MgGraph -Scopes "Application.Read.All"
    }
    else {
        $scopes = $context.Scopes
        if ('Application.Read.All' -notin $scopes) {
            Write-Host "⚠️ Current Graph session does not include required scope 'Application.Read.All'." -ForegroundColor Yellow
            Write-Host "🔁 Reconnecting with proper permissions..." -ForegroundColor Cyan
            Disconnect-MgGraph | Out-Null
            Connect-MgGraph -Scopes "Application.Read.All"
        }
        else {
            Write-Host "✅ Connected to Microsoft Graph with scope 'Application.Read.All'." -ForegroundColor Green
        }
    }

    Write-Host "🔍 Checking directory extensions..." -ForegroundColor Cyan
    Write-Host ""

    # This will store all extensions found
    $results = @()

    if ($AppDisplayName) {
        Write-Host "Looking up extensions for application: '$AppDisplayName'..." -ForegroundColor Cyan

        $app = Invoke-MgGraphRequest -Method GET "https://graph.microsoft.com/v1.0/applications?`$filter=displayName eq '$AppDisplayName'" |
        Select-Object -ExpandProperty value |
        Select-Object -ExpandProperty id
        $appID = $app.id

        if ($app) {
            $extensions = Invoke-MgGraphRequest -Method GET "https://graph.microsoft.com/v1.0/applications/$app/extensionProperties" |
            Select-Object -ExpandProperty value

            if ($extensions) {
                Write-Host "───────────────────────────────────────────────" -ForegroundColor DarkCyan
                Write-Host "✅ Found extensions for $AppDisplayName" -ForegroundColor Green
                Write-Host "───────────────────────────────────────────────" -ForegroundColor DarkCyan

                $extensions | ForEach-Object {
                    Write-Host ("Name: " + $_.name) -ForegroundColor Yellow
                    Write-Host ("ID: " + $_.id) -ForegroundColor Gray
                    Write-Host ("Targets: " + ($_.targetObjects -join ", ")) -ForegroundColor Magenta
                    Write-Host "-----------------------------------------------" -ForegroundColor DarkGray

                    $results += [PSCustomObject]@{
                        ApplicationName = $AppDisplayName
                        ApplicationID   = $appID
                        ExtensionName   = $_.name
                        ExtensionID     = $_.id
                        TargetObjects   = ($_.targetObjects -join ", ")
                    }
                }
            }
            else {
                Write-Host "⚠️ No extensions found for this app." -ForegroundColor DarkYellow
            }
        }
        else {
            Write-Host "❌ No application found with the name '$AppDisplayName'." -ForegroundColor Red
        }
    }
    else {
        Write-Host "No specific app provided – scanning all registered applications..." -ForegroundColor Yellow
        $url = "https://graph.microsoft.com/v1.0/applications"

        do {
            $response = Invoke-MgGraphRequest -Method GET -Uri $url
            $url = $response.'@odata.nextLink'

            $response.value | ForEach-Object {
                Write-Host ""
                Write-Host "→ Checking extensions for: $($_.displayName)" -ForegroundColor Cyan
                $AppDisplayName = "$($_.displayName)"
                $appID = $_.id

                $ext = Invoke-MgGraphRequest -Method GET "https://graph.microsoft.com/v1.0/applications/$($_.id)/extensionProperties"

                if ($ext.value.Count -gt 0) {
                    Write-Host "───────────────────────────────────────────────" -ForegroundColor DarkGreen
                    Write-Host "✅ Found $($ext.value.Count) extension(s) for $($_.displayName)" -ForegroundColor Green
                    Write-Host "───────────────────────────────────────────────" -ForegroundColor DarkGreen

                    $ext.value | ForEach-Object {
                        Write-Host ("Name: " + $_.name) -ForegroundColor Yellow
                        Write-Host ("ID: " + $_.id) -ForegroundColor Gray
                        Write-Host ("Targets: " + ($_.targetObjects -join ", ")) -ForegroundColor Magenta
                        Write-Host "-----------------------------------------------" -ForegroundColor DarkGray

                        $results += [PSCustomObject]@{
                            ApplicationName = $AppDisplayName
                            ApplicationID   = $appid
                            ExtensionName   = $_.name
                            ExtensionID     = $_.id
                            TargetObjects   = ($_.targetObjects -join ", ")
                        }
                    }
                }
                else {
                    Write-Host "⚠️ No extensions found." -ForegroundColor DarkYellow
                }
            }

        } while ($url)
    }

    Write-Host ""
    Write-Host "✅ Done listing directory extensions." -ForegroundColor Green

    # Return results for export or further use
    return $results

    <#
    .SYNOPSIS
    Lists all directory extensions on one or all registered Entra ID applications.

    .DESCRIPTION
    If you provide an application display name, it lists extensions for that app only.
    Otherwise, it loops through all registered applications and shows their extensions.

    .PARAMETER AppDisplayName
    The display name of the target application (optional).

    .EXAMPLE
    Get-DirectoryExtensions -AppDisplayName "MyApp"
    .EXAMPLE
    Get-DirectoryExtensions
    #>
}



function Get-DirectoryExtensionValues {
    param (
        [Parameter()]
        [String]$DirectoryExtensionName,
        [Parameter()]$UserUPN
    )

    #Connecting MgGraph 
       $context = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $context) {
        Write-Host "🔗 Not connected to Microsoft Graph. Connecting with required scope..." -ForegroundColor Cyan
        Connect-MgGraph -Scopes "User.Read.All"
    }
    else {
        $scopes = $context.Scopes
        if ('User.Read.All' -notin $scopes) {
            Write-Host "⚠️ Current Graph session does not include required scope 'User.Read.All'." -ForegroundColor Yellow
            Write-Host "🔁 Reconnecting with proper permissions..." -ForegroundColor Cyan
            Disconnect-MgGraph | Out-Null
            Connect-MgGraph -Scopes "User.Read.All"
        }
        else {
            Write-Host "✅ Connected to Microsoft Graph with scope 'User.Read.All'." -ForegroundColor Green
        }
    }

     $results = @()
    if ($UserUPN) {
        Write-Host "👤 Getting extension values for user $UserUPN..." -ForegroundColor Cyan
        $user = igall "https://graph.microsoft.com/beta/users/$UserUPN"
       if ($DirectoryExtensionName) {
            $value = $user.$DirectoryExtensionName
            if ($null -ne $value) {
                $results += [PSCustomObject]@{
                    DisplayName       = $user.displayName
                    UserPrincipalName = $user.userPrincipalName
                    ExtensionName     = $DirectoryExtensionName
                    ExtensionValue    = $value
                }
            }
        }
         else {
            $user.PSObject.Properties |
            Where-Object { $_.Name -like "extension_*" -and $null -ne $_.Value } |
            ForEach-Object {
                $results += [PSCustomObject]@{
                    DisplayName       = $user.displayName
                    UserPrincipalName = $user.userPrincipalName
                    ExtensionName     = $_.Name
                    ExtensionValue    = $_.Value
                }
            }
        }
    }
    else {
        Write-Host "📋 No user specified — retrieving directory extension values for all users (this may take a while)..." -ForegroundColor Yellow
       $users = Igall  "https://graph.microsoft.com/beta/users"
        
$users | Foreach-Object {
    $user = $_
    if ($DirectoryExtensionName) {
                    $value = $user.$DirectoryExtensionName
                    if ($null -ne $value) {
                        $results += [PSCustomObject]@{
                            DisplayName       = $user.displayName
                            UserPrincipalName = $user.userPrincipalName
                            ExtensionName     = $DirectoryExtensionName
                            ExtensionValue    = $value
                        }
                    }
                }
                else {
                    $user.PSObject.Properties |
                    Where-Object { $_.Name -like "extension_*" -and $null -ne $_.Value } |
                    ForEach-Object {
                        $results += [PSCustomObject]@{
                            DisplayName       = $user.displayName
                            UserPrincipalName = $user.userPrincipalName
                            ExtensionName     = $_.Name
                            ExtensionValue    = $_.Value
                        }
                    }
                }
            }
}


        

        Write-Host "✅ Finished collecting extension values." -ForegroundColor Green
        return $results

 <#
    .SYNOPSIS
    Fetches directory extension values for one or all users.

    .DESCRIPTION
    If a DirectoryExtensionName is provided, retrieves that specific value.
    If not, lists all extensions with non-null values for all users.

    .ROLE
    User Administrator

    .PARAMETER DirectoryExtensionName
    The specific directory extension to retrieve (optional).

    .PARAMETER UserUPN
    The UPN of a specific user (optional).
    #>

}
function Set-DirectoryExtensionValue {
    param (
        [Parameter(Mandatory)]
        [String]$DirectoryExtensionName,

        [Parameter(Mandatory)]
        [String]$UserUPN,

        [Parameter(Mandatory)]
        [String]$NewValue
    )

    #Connecting MgGraph 
    $context = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $context) {
        Write-Host "🔗 Not connected to Microsoft Graph. Connecting with required scope..." -ForegroundColor Cyan
        Connect-MgGraph -Scopes "Directory.ReadWrite.All"
    }
    else {
        $scopes = $context.Scopes
        if ('Directory.ReadWrite.All' -notin $scopes) {
            Write-Host "⚠️ Current Graph session does not include required scope 'Directory.ReadWrite.All'." -ForegroundColor Yellow
            Write-Host "🔁 Reconnecting with proper permissions..." -ForegroundColor Cyan
            Disconnect-MgGraph | Out-Null
            Connect-MgGraph -Scopes "Directory.ReadWrite.All"
        }
        else {
            Write-Host "✅ Connected to Microsoft Graph with scope 'Directory.ReadWrite.All'." -ForegroundColor Green
        }
    }

    # Prepare payload
    $payload = @{
        $DirectoryExtensionName = $NewValue
    } | ConvertTo-Json

    Write-Host "✏️ Setting '$DirectoryExtensionName' for user '$UserUPN' to '$NewValue'..." -ForegroundColor Cyan

    # Update user
    try {
        Invoke-MgGraphRequest -Method PATCH  -Uri "https://graph.microsoft.com/beta/users/$UserUPN"  -Body $payload  -ContentType "application/json"

        Write-Host "✅ Successfully updated $DirectoryExtensionName for $UserUPN." -ForegroundColor Green
    }
    catch {
        Write-Host "❌ Failed to update $DirectoryExtensionName for $UserUPN. Error:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor DarkRed
    }

    <#
.SYNOPSIS
Sets a specific directory extension value for a user in Microsoft Entra ID.

.DESCRIPTION
Updates the specified directory extension (custom schema attribute) for a given user.
Requires the Directory.ReadWrite.All permission scope.

.PARAMETER DirectoryExtensionName
The full name of the directory extension (e.g., extension_{AppClientId}_CustomAttribute).

.PARAMETER UserUPN
The UPN of the target user.

.PARAMETER NewValue
The new value to assign to the directory extension.

.EXAMPLE
Set-DirectoryExtensionValue -DirectoryExtensionName "extension_abc1234_responsibilities" -UserUPN "user@domain.com" -NewValue "Finance"
#>
}


function Remove-ApplicationDirectoryExtension {


    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [string[]]$ApplicationId,

        [Parameter(Mandatory)]
        [string]$ExtensionId
    )

    begin {
        #Connecting MgGraph 
    $context = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $context) {
        Write-Host "🔗 Not connected to Microsoft Graph. Connecting with required scope..." -ForegroundColor Cyan
        Connect-MgGraph -Scopes "Application.ReadWrite.All"
    }
    else {
        $scopes = $context.Scopes
        if ('Application.ReadWrite.All' -notin $scopes) {
            Write-Host "⚠️ Current Graph session does not include required scope 'Application.ReadWrite.All'." -ForegroundColor Yellow
            Write-Host "🔁 Reconnecting with proper permissions..." -ForegroundColor Cyan
            Disconnect-MgGraph | Out-Null
            Connect-MgGraph -Scopes "Application.ReadWrite.All"
        }
        else {
            Write-Host "✅ Connected to Microsoft Graph with scope 'Application.ReadWrite.All'." -ForegroundColor Green
        }
    }
}

    process {
        $ApplicationId | ForEach-Object {
            $appId = $_
            Write-Host "`nChecking for extension property on application: $appId" -ForegroundColor Yellow

            try {
                $extension = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/applications/$appId/extensionProperties/$ExtensionId" -ErrorAction Stop
            }
            catch {
                Write-Host "❌ No extension property found with ID '$ExtensionId' on application $appId." -ForegroundColor Red
                return
            }

            Write-Host "✅ Extension property found:" -ForegroundColor Green
            Write-Host "  Name: $($extension.name)"
            Write-Host "  ID: $($extension.id)"
            Write-Host "  Target App: $appId`n"

            $confirm = Read-Host "⚠️ This will remove the extension property for ALL users in the tenant that rely on it. Are you sure you want to continue? (Y/N)"
            if ($confirm -notin @('Y', 'y', 'Yes', 'yes')) {
                Write-Host "Operation cancelled for application $appId." -ForegroundColor Yellow
                return
            }

            try {
                Write-Host "Removing extension property..." -ForegroundColor Cyan
                Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/applications/$appId/extensionProperties/$ExtensionId" -ErrorAction Stop
                Write-Host "✅ Extension property successfully removed from application $appId." -ForegroundColor Green
            }
            catch {
                Write-Host "❌ Failed to remove extension property from $appId. Error: $_" -ForegroundColor Red
            }
        }
    }

    <#
    .SYNOPSIS
    Safely removes one or more extension properties from Microsoft Entra applications. 

    .DESCRIPTION
    This function checks if the specified extension property exists on one or more applications,
    confirms with the administrator before removal, and warns that the action affects all users
    in the tenant that rely on the extension.

    .PARAMETER ApplicationId
    One or more Object IDs or App IDs of applications in Microsoft Entra ID.
    Can be piped in.

    .PARAMETER ExtensionId
    The ID of the extension property to remove.

    .EXAMPLE
   Remove-ApplicationDirectoryExtension -ExtensionId 
    #>
}
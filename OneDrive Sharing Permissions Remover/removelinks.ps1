#Requires -Version 7.0
# Make sure to fill in all the required variables before running the script
# Also make sure the AppID used corresponds to an app with sufficient permissions, as follows:
#    User.Read.All to enumerate all users in the tenant
#    Sites.ReadWrite.All to return all the item sharing details
# Help file: https://github.com/michevnew/PowerShell/blob/master/Graph_ODFB_remove_all_shared.md
# More info at: https://www.michev.info/blog/post/3018/remove-sharing-permissions-on-all-files-in-users-onedrive-for-business

[CmdletBinding()] #Make sure we can use -Verbose
Param(
    [string]$TenantID ,
    [pscredential]$ClientSecretCredential,
    [switch]$ExpandFolders=$true,
    [int]$Depth = 2)

#==========================================================================
# Helper functions
#==========================================================================
function igall {
    [CmdletBinding()]
    param (
        [string]$Uri
    )
    $nextUri = $uri
    do {
        $result = Invoke-MgGraphRequest -Method GET -uri $nextUri
        $nextUri = $result.'@odata.nextLink'
        if ($result -and $result.ContainsKey('value')) {
            $result.value

        }
        else {
            $result
        }
    } while ($nextUri)
}
function processChildren {

    Param(
        #Graph User object
        [Parameter(Mandatory = $true)]$User,
        #URI for the drive
        [Parameter(Mandatory = $true)][string]$URI,
        #Use the ExpandFolders switch to specify whether to expand folders and include their items in the output.
        [switch]$ExpandFolders,
        #Use the Depth parameter to specify the folder depth for expansion/inclusion of items.
        [int]$depth)

    $URI = "$URI/children"
    $children = @()
    try {
        $children += igall $URI
    } catch {
    }

    if (!$children) { Write-Verbose "No child items found..."; return }

    #handle different children types
    $output = @()
    $cFolders = $children | Where-Object { $_.Folder }
    $cFiles = $children | Where-Object { $_.File } #doesnt return notebooks
    $cNotebooks = $children | Where-Object { $_.package.type -eq "OneNote" }

    #Process Folders
    foreach ($folder in $cFolders) {
        $output += (processFolder -User $User -folder $folder -ExpandFolders:$ExpandFolders -depth $depth -Verbose:$VerbosePreference)
    }

    #Process Files
    foreach ($file in $cFiles) {
        if ($file.shared) {
            Write-Host "Found shared file ($($file.name)), removing permissions..."
            RemovePermissions $User.id $file.id -Verbose:$VerbosePreference
            $fileinfo = New-Object psobject
            $fileinfo | Add-Member -MemberType NoteProperty -Name "Name" -Value $file.name
            $fileinfo | Add-Member -MemberType NoteProperty -Name "ItemType" -Value "File"
            $fileinfo | Add-Member -MemberType NoteProperty -Name "ItemPath" -Value $file.webUrl
            $output += $fileinfo
        }
        else { continue }
    }

    #Process Notebooks
    foreach ($notebook in $cNotebooks) {
        if ($notebook.shared) {
            Write-Host "Found shared notebook ($($notebook.name)), removing permissions..."
            RemovePermissions $User.id $notebook.id -Verbose:$VerbosePreference
            $fileinfo = New-Object psobject
            $fileinfo | Add-Member -MemberType NoteProperty -Name "Name" -Value $notebook.name
            $fileinfo | Add-Member -MemberType NoteProperty -Name "ItemType" -Value "Notebook"
            $fileinfo | Add-Member -MemberType NoteProperty -Name "ItemPath" -Value $notebook.webUrl
            $output += $fileinfo
        }
    }
    return $output
}

function processFolder {

    Param(
        #Graph User object
        [Parameter(Mandatory = $true)]$User,
        #Folder object
        [Parameter(Mandatory = $true)]$folder,
        #Use the ExpandFolders switch to specify whether to expand folders and include their items in the output.
        [switch]$ExpandFolders,
        #Use the Depth parameter to specify the folder depth for expansion/inclusion of items.
        [int]$depth)

    #if the Shared property is set, fetch permissions
    if ($folder.shared) {
        Write-Host "Found shared folder ($($folder.name)), removing permissions..."
        RemovePermissions $User.id $folder.id -Verbose:$VerbosePreference
        $fileinfo = New-Object psobject
        $fileinfo | Add-Member -MemberType NoteProperty -Name "Name" -Value $folder.name
        $fileinfo | Add-Member -MemberType NoteProperty -Name "ItemType" -Value "Folder"
        $fileinfo | Add-Member -MemberType NoteProperty -Name "ItemPath" -Value $folder.webUrl
    }

    #Since this is a folder item, check for any children, depending on the script parameters
    if (($folder.folder.childCount -gt 0) -and $ExpandFolders -and ((3 - $folder.parentReference.path.Split("/").Count + $depth) -gt 0)) {
        Write-Verbose "Folder $($folder.Name) has child items"
        $uri = "https://graph.microsoft.com/v1.0/users/$($user.id)/drive/items/$($folder.id)"
        $folderItems = processChildren -User $user -URI $uri -ExpandFolders:$ExpandFolders -depth $depth -Verbose:$VerbosePreference
    }

    #handle the output
    if ($folderItems) { $f = @(); $f += $fileinfo; $f += $folderItems; return $f }
    else { return $fileinfo }
}

function RemovePermissions {

    Param(
        #Use the UserId parameter to provide an unique identifier for the user object.
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$UserId,
        #Use the ItemId parameter to provide an unique identifier for the item object.
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$ItemId)

    #fetch permissions for the given item
    $permissions = @()
    $uri = "https://graph.microsoft.com/beta/users/$($UserId)/drive/items/$($ItemId)/permissions"
    $permissions = igall $uri
    if (-not $permissions) {
        continue
    }

    foreach ($entry in $permissions) {
        if ($entry.inheritedFrom) { Write-Verbose "Skipping inherited permissions..." ; continue }
        Write-Host -ForegroundColor Green "Tar bort $($entry.id)"
        Invoke-MgGraphRequest -Method DELETE -Verbose:$VerbosePreference -Uri "$uri/$($entry.id)" -Headers $authHeader -SkipHeaderValidation -ErrorAction Stop | Out-Null
    }
    #check for sp. prefix on permission entries
    #SC admin permissions are skipped, not covered via the "shared" property
}

#==========================================================================
# Main script starts here
#==========================================================================

Connect-MgGraph -TenantId $TenantID -ClientSecretCredential $ClientSecretCredential

#prepare auth header
$global:authHeader = @{
    'Content-Type' = 'application\json'
}

#Check the user object
Write-Verbose "Checking user $user ..."
igall 'https://graph.microsoft.com/v1.0/users/' | foreach-object {
    $user = $_.userPrincipalName
    Write-Host "Processing user $user ODFB drive..."
    #Check whether the user has ODFB drive provisioned
    $uri = "https://graph.microsoft.com/v1.0/users/$($_.id)/drive/root"
    try { $UserDrive = Invoke-MgGraphRequest -Uri $uri -Verbose:$VerbosePreference -Headers $authHeader -ErrorAction Stop }
    catch { Write-Warning "User $user doesn't have OneDrive provisioned, ignoring..." }


    #If no items in the drive, skip
    if ( (-not $UserDrive) -or ($UserDrive.folder.childCount -eq 0)) {
        Write-Host "No items found for user $user"
    } else {
        #enumerate items in the drive and prepare the output
        Write-Verbose "Processing drive items..."
        processChildren -User $_ -URI $uri -ExpandFolders:$true -depth $depth
    }
}
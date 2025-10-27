
#helper method to manager fething paged results
function igall {
    [CmdletBinding()]
    param (
        [string]$Uri
    )
    $nextUri = $uri
    do {
        $result = Invoke-MgGraphRequest -Method GET -Uri $nextUri
        $nextUri = $result.'@odata.nextLink'
        if ($result -and $result.ContainsKey('value')) {
            $result.value
        }
        else {
            $result
        }
    } while ($nextUri)
}



$cache = @{}
function Get-User {
    param (
        [String]$Id
    )
    if (-not $cache[$id]) {
        $cache[$id] = igall "https://graph.microsoft.com/beta/users/$($id)?`$select=Displayname%2CUserprincipalname%2CcompanyName%2CaccountEnabled%2CCreatedDatetime%2CLastPasswordChangeDateTime%2csignInActivity%2clastNonInteractiveSignInDateTime%2clastSignInDateTime"
    }
    return [pscustomobject]$cache[$id]
}

function test-module {
    [CmdletBinding()]
    param(
        [String]$Name
  
    )
    Write-Host "Checking module $name..." -ForegroundColor Cyan
    if (-not (Get-Module $Name)) {
        Write-Host "Module $Name not imported, trying to import..." -ForegroundColor Yellow
        try {
            if ($Name -eq 'Microsoft.Graph') {
                Write-Host "Microsoft.Graph module import takes a while..." -ForegroundColor Yellow
                Import-Module $Name  -ErrorAction Stop
            }
            elseif ($Name -eq 'Az') {
                Write-Host "Module Az is being imported. This might take a while..." -ForegroundColor Yellow
            }
            else {
                Import-Module $Name  -ErrorAction Stop
            }
        
        }
        catch {
            Write-Host "Module $Name not found, trying to install..." -ForegroundColor Yellow
            Install-Module $Name -Scope CurrentUser -AllowClobber -Force -AcceptLicense -SkipPublisherCheck
            Write-Host "Importing module $Name..." -ForegroundColor Yellow
            Import-Module $Name -ErrorAction Stop
        }
    } 
    else {
        Write-Host "Module $Name is already imported." -ForegroundColor Green
    }   
}
#imports /installs require modules 
Write-Host "Importing required modules: Az.Resources, Az.Accounts, Microsoft.Graph.Authentication, ImportExcel" -ForegroundColor Cyan
Test-Module -Name Az.Resources
Test-Module -Name Az.Accounts
Test-Module -Name Microsoft.Graph.Authentication
Test-Module -Name ImportExcel
Write-Host "✅ All modules are installed and imported." -ForegroundColor Green
# Disconnect any existing sessions
Write-Host "Disconnecting any existing sessions..." -ForegroundColor Cyan
disconnect-Mggraph
Disconnect-AzAccount  
Write-Host "✅ All sessions disconnected." -ForegroundColor Green
# Connect new sessions
Write-Host "Connecting to AzAccount..." -ForegroundColor Cyan
Connect-AzAccount
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes 'RoleManagement.Read.Directory', 'User.Read.All', 'User.ReadBasic.All', 'User.Read', 'GroupMember.Read.All', 'Group.Read.All', 'Directory.Read.All', 'Directory.AccessAsUser.All', 'RoleEligibilitySchedule.Read.Directory', 'RoleManagement.Read.All', 'SecurityActions.Read.All', 'SecurityActions.ReadWrite.All', 'SecurityEvents.Read.All', "Organization.Read.All", "AuditLog.Read.All" -TenantId $tenantID  -ContextScope Process
Write-Host "✅ You are now fully connected!" -ForegroundColor Green


# Select folder for export 
Write-Host "--------------------------------------------------------" -ForegroundColor DarkGray
Write-Host "Please select a folder where the report will be saved." -ForegroundColor Cyan
Write-Host "⚠️  The folder selection window may appear behind other open windows." -ForegroundColor Yellow
Write-Host "If you don't see it, try minimizing other windows." -ForegroundColor Yellow
Write-Host "--------------------------------------------------------" -ForegroundColor DarkGray
Add-Type -AssemblyName System.Windows.Forms
$FileBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop') 
}
$result = $FileBrowser.ShowDialog(((New-Object System.Windows.Forms.Form -Property @{TopMost = $true })))
if ($result -eq [Windows.Forms.DialogResult]::OK) {
    $folder = $FileBrowser.SelectedPath
    Write-Host "Export folder selected: $folder" -ForegroundColor Green
}
else {
    Write-Host "❌ No folder selected. Exiting script." -ForegroundColor Red
    return
}


#Get org displayname
Write-Host "Fetching organization display name..." -ForegroundColor Yellow
$orgdisplayname = igall https://graph.microsoft.com/beta/organization | Select-Object -ExpandProperty displayName
Write-Host "Organization: $orgdisplayname" -ForegroundColor Green
    
$date = Get-Date -Format yyyy-MM-dd
Write-Host "Fetching directory roles..." -ForegroundColor Yellow
$directoryRoles = igall https://graph.microsoft.com/beta/directoryRoles | foreach-object {
    [PsCustomObject]$_
}
Write-Host "✅ Retrieved $($directoryRoles.Count) directory roles." -ForegroundColor Green

# In a live tenant the will be a lot of instances so we filter
# on endDateTime to limit the responses to active instances
$now = (Get-Date -AsUTC).ToString("yyyy-MM-ddTHH:mm:ssZ")
Write-Host "Fetching active role assignments (PIM activated roles)..." -ForegroundColor Yellow
$assignmentSchedules = @()
$assignmentSchedules += igall  "https://graph.microsoft.com/beta/roleManagement/directory/roleAssignmentScheduleInstances?`$expand=roleDefinition,principal&`$filter=assignmentType eq 'Activated' and endDateTime ge $now" | ForEach-Object {
    [PsCustomObject]$_
} |
Where-Object {
    $_.RoleDefinitionId -in $directoryRoles.roleTemplateId
}
Write-Host "✅ Retrieved $($assignmentSchedules.Count) active assignment schedules." -ForegroundColor Green
# Fetch role assignments to be able to filter out
# admins that have used PIM to activate a role
Write-Host "Fetching assignment admins..." -ForegroundColor Yellow
$assignmentAdmins = @()
$assignmentAdmins += $assignmentSchedules | Where-Object {
    $_.principal.'@odata.type' -notmatch "#microsoft.graph.servicePrincipal"
} | ForEach-Object {
    $assignment = $_
    if ($_.principal.'@odata.type' -match '#microsoft.graph.group') {
        igall -Uri "https://graph.microsoft.com/beta/groups/$($assignment.principalId)/transitiveMembers" | ForEach-Object {
            $member = [pscustomobject]$_
            Add-Member -InputObject $member -NotePropertyName 'Role' -NotePropertyValue $assignment.roleDefinition.displayName -PassThru
        }
    }
    else {
        $member = Get-User -Id $_.principalId
        Add-Member -InputObject $member -NotePropertyName 'Role' -NotePropertyValue $assignment.roleDefinition.displayName -PassThru
    }
}
Write-Host "✅ Assignment admins processed." -ForegroundColor Green
Write-Host "Building administrator list..." -ForegroundColor Yellow
$administrators = $directoryRoles | ForEach-Object {
    $role = $_.displayName    
    Write-Host "───────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "🔹 Processing directory role: $role" -ForegroundColor Cyan
    Write-Host " → Fetching members of role '$role'..." -ForegroundColor DarkGray
    igall -Uri "https://graph.microsoft.com/beta/directoryRoles/$($_.id)/members" | ForEach-Object {
        $member = [PSCustomObject]$_
        if ($member.'@odata.type' -notmatch 'group|ServicePrincipal') {
            Write-Host "   ↳ Found user: $($member.displayName)" -ForegroundColor Cyan
            Write-Host "     → Getting user details from Graph..." -ForegroundColor DarkGray
            $member = Get-User -id $member.id 
            Write-Host "     → Adding user '$($member.DisplayName)' to role '$role'" -ForegroundColor Yellow
            Add-Member -InputObject $member -NotePropertyName 'Role' -NotePropertyValue $role -PassThru -Force |
            Add-Member -NotePropertyName 'lastSignInDateTime' -NotePropertyValue $member.signInActivity.lastSignInDateTime -PassThru -Force
            Write-Host "     ✅ Completed: $($member.DisplayName)" -ForegroundColor Green
        }
        elseif ($member.'@odata.type' -match 'group') {
            Write-Host "   ↳ Expanding group: $($member.displayName)" -ForegroundColor Cyan
            Write-Host "     → Fetching transitive members..." -ForegroundColor DarkGray

            igall -Uri "https://graph.microsoft.com/beta/groups/$($member.id)/transitiveMembers" | ForEach-Object {
                Write-Host "       ↳ Adding group member: $($_.displayName)" -ForegroundColor Yellow
                $member = [PSCustomObject]$_
                Add-Member -InputObject $member -NotePropertyName 'Role' -NotePropertyValue $role -PassThru  
                Write-Host "       ✅ Added $($groupMember.DisplayName) (from group $($member.displayName))" -ForegroundColor Green
            }
            
        }
    }
} | Where-Object {
    # Filter out PIM activated admins as they
    # are displayed in the Eligible sheet
    $admin = $_
    $foundInAssignments = $assignmentAdmins | Where-Object {
        $admin.id -match $_.id -and $admin.Role -match $_.Role
    }
    -not $foundInAssignments
} | Select-Object Role, Displayname, Userprincipalname, companyName, accountEnabled, CreatedDatetime , LastPasswordChangeDateTime, lastSignInDateTime
Write-Host "✅ Administrator list compiled successfully." -ForegroundColor Green
Write-Host "───────────────────────────────────────────────" -ForegroundColor DarkGray
Write-Host "Fetching eligible roles..." -ForegroundColor Yellow
$eligible = igall -Uri 'https://graph.microsoft.com/beta/roleManagement/directory/roleEligibilityScheduleInstances/?$expand=roleDefinition,principal' | ForEach-Object {
    $e = [PSCustomObject]$_
    $principal = [PSCustomObject]$e.principal
        
    if ($e.memberType -match 'Direct' -and $principal.'@odata.type' -notmatch 'group|ServicePrincipal') {
        Write-Host "Processing eligible direct user: $($principal.displayName)" -ForegroundColor Cyan
        Write-Host " → Fetching detailed info for $($principal.userPrincipalName)" -ForegroundColor DarkGray
        $principal = Get-User -id $principal.id  | Select-Object Displayname, Userprincipalname, companyName, accountEnabled, CreatedDatetime , LastPasswordChangeDateTime, signInActivity
        Write-Host " → Adding role '$($e.roleDefinition["displayName"])' (MemberType: $($e.memberType))" -ForegroundColor Yellow
            
        Add-Member -InputObject $principal -NotePropertyName "EligibleRole" -NotePropertyValue $e.roleDefinition["displayName"] -PassThru |
        Add-Member -NotePropertyName 'Membertype' -NotePropertyValue $e.membertype -PassThru |
        Add-Member -NotePropertyName 'lastSignInDateTime' -NotePropertyValue $principal.signInActivity.lastSignInDateTime -PassThru -Force
        Write-Host " ✅ Completed processing for $($principal.DisplayName)" -ForegroundColor Green

    }
    elseif ($principal.'@odata.type' -match 'group') {
        Write-Host "Expanding eligible group: $($e.principal.displayName)" -ForegroundColor Cyan
        Write-Host " → Fetching members from group ID: $($e.principalId)" -ForegroundColor DarkGray
        $groupMembers = igall -Uri "https://graph.microsoft.com/beta/groups/$($e.principalId)/transitiveMembers" 
        $total = $groupMembers.Count
        $counter = 0
        $groupMembers | ForEach-Object  -Begin {
            Write-Progress -Activity "Expanding group: $($e.principal.displayName)" -Status "0 of $total members" -PercentComplete 0
        } -Process {
            $counter++
            $percent = [math]::Round(($counter / $total) * 100, 2)
            Write-Progress -Activity "Expanding group: $($e.principal.displayName)" -Status "$counter of $total members" -PercentComplete $percent
            $member = Get-User -id $_.id  | Select-Object Displayname, Userprincipalname, companyName, accountEnabled, CreatedDatetime , LastPasswordChangeDateTime, signInActivity
            Add-Member -InputObject $member -NotePropertyName "EligibleRole" -NotePropertyValue $e.roleDefinition["displayName"]-PassThru |
            Add-Member -NotePropertyName 'lastSignInDateTime' -NotePropertyValue $member.signInActivity.lastSignInDateTime -PassThru |
            Add-Member -NotePropertyName "EligibleRoleGroup" -NotePropertyValue $e.principal["displayName"] -PassThru -Force 
            Write-Host "     ✅ Added $($member.DisplayName) from group $($e.principal.displayName)" -ForegroundColor Green
        } -End {
        Write-Progress -Activity "Expanding group: $($e.principal.displayName)" -Completed
        Write-Host "✅ Finished expanding group $($e.principal.displayName) ($total members)" -ForegroundColor Green
    }
    }
} |  Select-Object displayName, Userprincipalname, EligibleRole, DirectRole, EligibleRoleGroup, memberType, createdDateTime, LastPasswordChangeDateTime, lastSignInDateTime
Write-Host "✅ Finished collecting all eligible role assignments." -ForegroundColor Green
Write-Host "Fetching Azure role assignments..." -ForegroundColor Yellow
$azroles = Get-AzSubscription | ForEach-Object {
    $id = $_.id 
    $name = $_.name 
    Write-Host "Fetching role assignments for subscription: $name" -ForegroundColor DarkCyan
    Get-AzRoleAssignment -Scope /subscriptions/$id | ForEach-Object {
        Add-Member -InputObject $_ -NotePropertyName 'Subscription' -NotePropertyValue $name -PassThru
    }
    

} | Select-Object roleDefinitionName, Displayname, SigninName, ObjectId, ObjectType, Subscription 
Write-Host "✅ Azure role assignments gathered." -ForegroundColor Green
Write-Host "Exporting data to Excel..." -ForegroundColor Cyan
$exportPath = "$folder\$orgdisplayname-EntraIDAdminReport$date.xlsx"
$administrators | Export-Excel -Path "$exportPath" -WorksheetName "Administrators" -StartRow 2 -TableName Adminstrators -AutoSize
$eligible | Export-Excel -Path "$exportPath" -WorksheetName "Eligible Roles" -StartRow 2 -TableName Adminbygroups -AutoSize
$azroles | Export-Excel -Path "$exportPath" -WorksheetName "Azure Roles" -StartRow 2 -TableName azAdmins -AutoSize
Write-Host "✅ Export completed successfully: $exportPath" -ForegroundColor Green

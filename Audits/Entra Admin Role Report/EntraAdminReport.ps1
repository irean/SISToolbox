[CmdletBinding()]
param()
#helper method to manager fething paged results
function igall {
    [CmdletBinding()]
    param (
        [string]$Uri
    )
    $nextUri = $uri
    do {
        $result = $null
        $time = Measure-Command { 
            $result = Invoke-MgGraphRequest -Method GET -Uri $nextUri
        }
        Write-Debug "callto $nextURI took $time"
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
        [string]$Id
    )

    
    if (-not $cache[$Id]) {

        $user = igall "https://graph.microsoft.com/v1.0/users/$($id)?`$select=Displayname%2CUserprincipalname%2CcompanyName%2CaccountEnabled%2CCreatedDatetime%2CLastPasswordChangeDateTime%2csignInActivity%2clastNonInteractiveSignInDateTime%2clastSignInDateTime%2CassignedLicenses%2CassignedPlans"

        $result = [pscustomobject]$user
        $result | Add-Member -NotePropertyName lastSignInDateTime -NotePropertyValue $user.signInActivity.lastSignInDateTime -Force
        $result | Add-Member -NotePropertyName lastNonInteractiveSignInDateTime -NotePropertyValue $user.signInActivity.lastNonInteractiveSignInDateTime -Force
        $result | Add-Member -NotePropertyName hasStrongMFA -NotePropertyValue $false -Force

        Start-Sleep -Milliseconds 250 

        $auth = Invoke-MgGraphRequest -Method GET  -Uri "https://graph.microsoft.com/beta/users/$Id/authentication/methods"
        $count = $auth.value.'@odata.type' | Where-Object {
            $_ -notmatch 'passwordAuthenticationMethod|phoneAuthenticationMethod'
        } | measure-object 

        $result | Add-Member -NotePropertyName StrongAuthCount -NotePropertyValue $count.count -Force

        foreach ($method in $auth.value) {

            switch ($method.'@odata.type') {

                '#microsoft.graph.passwordAuthenticationMethod' {
                    $result | Add-Member -NotePropertyName AuthPassword -NotePropertyValue $method.createdDateTime -Force

                }

                '#microsoft.graph.phoneAuthenticationMethod' {
                    $result | Add-Member -NotePropertyName AuthPhone -NotePropertyValue $method.phoneNumber -Force

                    
                }

                '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod' {
                    $result | Add-Member -NotePropertyName AuthMicrosoftAuthenticator -NotePropertyValue $method.displayName -Force
                    $result.hasStrongMFA = $true
                }

                '#microsoft.graph.passwordlessMicrosoftAuthenticatorAuthenticationMethod' {
                    $result | Add-Member -NotePropertyName AuthPasswordless -NotePropertyValue $method.displayName -Force
                    $result.hasStrongMFA = $true
                }

                '#microsoft.graph.fido2AuthenticationMethod' {

                    $result | Add-Member -NotePropertyName AuthFido2 -NotePropertyValue $method.displayName -Force
                    $result.hasStrongMFA = $true

                }
            }
        }
        if ($user.assignedLicenses.Count -gt 0) {
            $result | Add-Member -NotePropertyName IsLicensed -NotePropertyValue $true -Force
        }
        else {
            $result | Add-Member -NotePropertyName IsLicensed -NotePropertyValue $false -Force
        }


        $enabledProductivityPlans = $user.assignedPlans | Where-Object {
            $_.capabilityStatus -eq "Enabled" -and
            $_.service -match "EXCHANGE|SHAREPOINT|TEAMS|FLOW|POWERAPPS"
        }

        if ($enabledProductivityPlans) {
            $result | Add-Member -NotePropertyName ProductivityServicesEnabled -NotePropertyValue $true -Force
        }
        else {
            $result | Add-Member -NotePropertyName ProductivityServicesEnabled -NotePropertyValue $false -Force
        }

        $productivityServices = $enabledProductivityPlans.service |
        Sort-Object -Unique

        $result | Add-Member -NotePropertyName ProductivityServices `
            -NotePropertyValue ($productivityServices -join ", ") `
            -Force

 

        $cache[$Id] = $result
    }

    return $cache[$Id]
}

function Get-AdminRiskScore {
    param(
        $User,
        $Role
    )

    $score = 0

    # Disabled accounts = no risk
    if ($User.accountEnabled -eq $false) {
        return 0
    }

    # ROLE RISK
    $RoleRiskTable = @{
        "Global Administrator"          = 10
        "Privileged Role Administrator" = 9
        "Security Administrator"        = 8
        "User Administrator"            = 7
        "Groups Administrator"          = 6
    }
    if ($RoleRiskTable.ContainsKey($Role)) {
        $score += $RoleRiskTable[$Role]
    }
    else {
        $score += 3
    }

    # MFA RISK
    if (-not $User.hasStrongMFA) {
        $score += 10
    }
    elseif ($User.StrongAuthCount -lt 2) {
        $score += 2
    }

    # INACTIVE ADMIN
    if ($User.lastSignInDateTime) {
        $last = [datetime]$User.lastSignInDateTime
        if ($last -lt (Get-Date).AddDays(-90)) {
            $score += 3
        }
    }

    # PRODUCTIVITY SERVICES
    if ($User.ProductivityServicesEnabled) {
        $score += 5
    }
    if ($User.IsLicensed -and -not $User.hasStrongMFA) {
        $score += 5
    }
    if ($User.LastPasswordChangeDateTime) {
        if ([datetime]$User.LastPasswordChangeDateTime -lt (Get-Date).AddDays(-365)) {
            $score += 2
        }
    }
    # Service Principals
    if ($User.ObjectType -eq "ServicePrincipal" -and $score -ge 8) {
        $score += 5
    }

    return $score
}
function Get-AdminRiskLevel {

    param($Score)

    if ($Score -ge 15) { return "Critical" }
    elseif ($Score -ge 10) { return "High" }
    elseif ($Score -ge 5) { return "Medium" }
    else { return "Low" }
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
Disable-AzContextAutosave -Scope Process
Update-AzConfig -LoginExperienceV2 Off -Scope Process
Connect-AzAccount
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes 'RoleManagement.Read.Directory', 'User.Read.All', 'User.ReadBasic.All', 'User.Read', 'GroupMember.Read.All', 'Group.Read.All', 'Directory.Read.All', 'Directory.AccessAsUser.All', 'RoleEligibilitySchedule.Read.Directory', 'RoleManagement.Read.All', 'SecurityActions.Read.All', 'SecurityActions.ReadWrite.All', 'SecurityEvents.Read.All', "Organization.Read.All", "AuditLog.Read.All"   -ContextScope Process
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
Write-Host "Caching All Users" -ForegroundColor Cyan

# igall "https://graph.microsoft.com/beta/users?`$select=Displayname%2CUserprincipalname%2CcompanyName%2CaccountEnabled%2CCreatedDatetime%2CLastPasswordChangeDateTime%2csignInActivity%2clastNonInteractiveSignInDateTime%2clastSignInDateTime" | Foreach-Object {
#     $cache.add($_.id, $_)
# }
Write-Host "✅ Retrieved all users" -ForegroundColor Green
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
        $member = $member | Select-Object *
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
            $member = $member | Select-Object *

            Write-Host "     → Adding user '$($member.DisplayName)' to role '$role'" -ForegroundColor Yellow

            $member = Add-Member -InputObject $member -NotePropertyName 'Role' -NotePropertyValue $role -PassThru -Force

            $member = Add-Member -InputObject $member -NotePropertyName 'lastSignInDateTime' -NotePropertyValue $member.signInActivity.lastSignInDateTime -PassThru -Force
            $riskScore = Get-AdminRiskScore -User $member -Role $role
            $riskLevel = Get-AdminRiskLevel $riskScore

            $member | Add-Member -NotePropertyName 'AdminRiskScore' -NotePropertyValue $riskScore -Force
            $member | Add-Member -NotePropertyName 'AdminRiskLevel' -NotePropertyValue $riskLevel -Force        
            Write-Host "     ✅ Completed: $($member.DisplayName)" -ForegroundColor Green
            $member
        }
        elseif ($member.'@odata.type' -match 'group') {
            Write-Host "   ↳ Expanding group: $($member.displayName)" -ForegroundColor Cyan
            Write-Host "     → Fetching transitive members..." -ForegroundColor DarkGray

            igall -Uri "https://graph.microsoft.com/beta/groups/$($member.id)/transitiveMembers" | ForEach-Object {
                Write-Host "       ↳ Adding group member: $($_.displayName)" -ForegroundColor Yellow
                $member = [PSCustomObject]$_
                Add-Member -InputObject $member -NotePropertyName 'Role' -NotePropertyValue $role -PassThru  
                Write-Host "       ✅ Added $($member.DisplayName) (from group $($member.displayName))" -ForegroundColor Green
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
} | Select-Object Role, displayName, Userprincipalname, companyName, AdminRiskScore, AdminRiskLevel, accountEnabled, CreatedDatetime , LastPasswordChangeDateTime, lastSignInDateTime, hasStrongMFA, StrongAuthCount, AuthPassword, AuthPhone, AuthFido2, AuthPasswordless, AuthMicrosoftAuthenticator, IsLicensed

Write-Host "✅ Administrator list compiled successfully." -ForegroundColor Green
Write-Host "───────────────────────────────────────────────" -ForegroundColor DarkGray
Write-Host "Fetching eligible roles..." -ForegroundColor Yellow
$eligible = igall -Uri 'https://graph.microsoft.com/beta/roleManagement/directory/roleEligibilityScheduleInstances/?$expand=roleDefinition,principal' | ForEach-Object {
    $e = [PSCustomObject]$_
    $principal = [PSCustomObject]$e.principal
        
    if ($e.memberType -match 'Direct' -and $principal.'@odata.type' -notmatch 'group|ServicePrincipal') {
        Write-Host "Processing eligible direct user: $($principal.displayName)" -ForegroundColor Cyan
        Write-Host " → Fetching detailed info for $($principal.userPrincipalName)" -ForegroundColor DarkGray
        $principal = Get-User -id $principal.id  
        Write-Host " → Adding role '$($e.roleDefinition["displayName"])' (MemberType: $($e.memberType))" -ForegroundColor Yellow
            
        $principal | Add-Member -NotePropertyName "EligibleRole" -NotePropertyValue $e.roleDefinition.displayName -Force
        $principal | Add-Member -NotePropertyName "MemberType" -NotePropertyValue $e.memberType -Force

        if ($principal.signInActivity) {
            $principal | Add-Member lastSignInDateTime $principal.signInActivity.lastSignInDateTime -Force
        }
        $riskScore = Get-AdminRiskScore -User $principal -Role $principal.EligibleRole
        $riskLevel = Get-AdminRiskLevel $riskScore

        $principal | Add-Member AdminRiskScore $riskScore -Force
        $principal | Add-Member AdminRiskLevel $riskLevel -Force

        $principal
        Write-Host " ✅ Completed processing for $($principal.DisplayName)" -ForegroundColor Green

    }
    elseif ($principal.'@odata.type' -match 'group') {
        Write-Host "Expanding eligible group: $($e.principal.displayName)" -ForegroundColor Cyan
        Write-Host " → Fetching members from group ID: $($e.principalId)" -ForegroundColor DarkGray
        $groupMembers = igall -Uri "https://graph.microsoft.com/beta/groups/$($e.principalId)/transitiveMembers" | Sort-Object id -Unique
        $total = $groupMembers.Count
        $counter = 0
        $groupMembers | ForEach-Object  -Begin {
            Write-Progress -Activity "Expanding group: $($e.principal.displayName)" -Status "0 of $total members" -PercentComplete 0
        } -Process {
            $counter++
            $percent = [math]::Round(($counter / $total) * 100, 2)
            Write-Progress -Activity "Expanding group: $($e.principal.displayName)" -Status "$counter of $total members" -PercentComplete $percent
            if ($_.'@odata.type' -eq '#microsoft.graph.user') {

                $member = Get-User -id $_.id
                $member = $member | Select-Object *
                $member | Add-Member IdentityType "User" -Force


            }
            elseif ($_.'@odata.type' -eq '#microsoft.graph.servicePrincipal') {

                $member = [pscustomobject]@{
                    displayName       = $_.displayName
                    UserPrincipalName = $null
                    IdentityType      = "ServicePrincipal"
                    hasStrongMFA      = $false
                    StrongAuthCount   = 0
                }

            }
            elseif ($_.'@odata.type' -eq '#microsoft.graph.group') {

                Write-Warning "Nested group detected: $($_.displayName)"

                $nested = [pscustomobject]@{
                    displayName                = $_.displayName
                    UserPrincipalName          = $null
                    EligibleRole               = $e.roleDefinition.displayName
                    DirectRole                 = $null
                    EligibleRoleGroup          = $e.principal.displayName
                    AdminRiskScore             = 7
                    AdminRiskLevel             = "Medium"
                    memberType                 = "NestedGroup"
                    createdDateTime            = $null
                    LastPasswordChangeDateTime = $null
                    lastSignInDateTime         = $null
                    hasStrongMFA               = $false
                    StrongAuthCount            = 0
                    AuthPassword               = $null
                    AuthPhone                  = $null
                    AuthFido2                  = $null
                    AuthPasswordless           = $null
                    AuthMicrosoftAuthenticator = $null
                }

                $nested

                continue
            }
            $member = $member | Select-Object *
            Add-Member -InputObject $member -NotePropertyName "EligibleRole" -NotePropertyValue $e.roleDefinition["displayName"] -PassThru -Force |

            Add-Member -NotePropertyName "MemberType" -NotePropertyValue "Group" -PassThru -Force |

            Add-Member -NotePropertyName "EligibleRoleGroup" -NotePropertyValue $e.principal.displayName -PassThru -Force
            if ($member.signInActivity) {
                $member | Add-Member lastSignInDateTime $member.signInActivity.lastSignInDateTime -Force
            }
            $riskScore = Get-AdminRiskScore -User $member -Role $member.EligibleRole
            $riskLevel = Get-AdminRiskLevel $riskScore

            $member | Add-Member -NotePropertyName AdminRiskScore -NotePropertyValue $riskScore -Force
            $member | Add-Member -NotePropertyName AdminRiskLevel -NotePropertyValue $riskLevel -Force

            $member
            Write-Host "     ✅ Added $($member.DisplayName) from group $($e.principal.displayName)" -ForegroundColor Green
        } -End {
            Write-Progress -Activity "Expanding group: $($e.principal.displayName)" -Completed
            Write-Host "✅ Finished expanding group $($e.principal.displayName) ($total members)" -ForegroundColor Green
        }
    }
} |  Select-Object id, displayName, Userprincipalname, EligibleRole, DirectRole, EligibleRoleGroup, memberType, AdminRiskScore, AdminRiskLevel, createdDateTime, LastPasswordChangeDateTime, lastSignInDateTime, hasStrongMFA, StrongAuthCount, AuthPassword, AuthPhone, AuthFido2, AuthPasswordless, AuthMicrosoftAuthenticator, IsLicensed
Write-Host "✅ Finished collecting all eligible role assignments." -ForegroundColor Green
Write-Host "Fetching Azure role assignments..." -ForegroundColor Yellow
$azroles = Get-AzSubscription | ForEach-Object {
    $id = $_.id 
    $name = $_.name 

    Write-Host "Fetching role assignments for subscription: $name" -ForegroundColor DarkCyan

    Get-AzRoleAssignment -Scope /subscriptions/$id | ForEach-Object {

        $assignment = $_
        $assignmentSource = "Direct"

        if ($assignment.ObjectType -eq "Group") {
            $assignmentSource = "Group"
        }
        $isLicensed = $null
        $productivityEnabled = $null

        if ($assignment.ObjectType -eq "User") {
            $user = Get-User -Id $assignment.ObjectId
            $isLicensed = $user.IsLicensed
            $productivityEnabled = $user.ProductivityServicesEnabled
        }



        $riskScore = 0

        if ($assignment.RoleDefinitionName -match "Owner") { $riskScore = 10 }
        elseif ($assignment.RoleDefinitionName -match "User Access Administrator") { $riskScore = 8 }
        elseif ($assignment.RoleDefinitionName -match "Contributor") { $riskScore = 6 }
        elseif ($assignment.RoleDefinitionName -match "Reader") { $riskScore = 1 }

        $riskLevel = Get-AdminRiskLevel $riskScore

        $assignment | Add-Member -NotePropertyName 'Subscription' -NotePropertyValue $name -Force
        $assignment | Add-Member -NotePropertyName 'AdminRiskScore' -NotePropertyValue $riskScore -Force
        $assignment | Add-Member -NotePropertyName 'AdminRiskLevel' -NotePropertyValue $riskLevel -Force
        $assignment | Add-Member -NotePropertyName 'IsLicensed' -NotePropertyValue $isLicensed -Force
        $assignment | Add-Member -NotePropertyName 'AssignmentSource' -NotePropertyValue $assignmentSource -Force
        $assignment | Add-Member -NotePropertyName 'ProductivityServicesEnabled' -NotePropertyValue $productivityEnabled -Force

        $assignment
    }

} | Select-Object roleDefinitionName, displayname, SigninName, ObjectId, ObjectType, AssignmentSource, Subscription, AdminRiskScore, AdminRiskLevel, IsLicensed, ProductivityServicesEnabled





Write-Host "✅ Azure role assignments gathered." -ForegroundColor Green

$exportPath = "$folder\$orgdisplayname-EntraIDAdminReport$date.xlsx"
if (Test-Path $exportPath) {
    Write-Host "Removing existing report..." -ForegroundColor Yellow
    Remove-Item $exportPath -Force
}
$allAdmins = @()

# -------------------------
# Administrators
# -------------------------
$allAdmins += $administrators | ForEach-Object {
    [PSCustomObject]@{
        UserPrincipalName           = $_.UserPrincipalName
        ObjectId                    = $_.Id
        displayName                 = $_.displayName
        Role                        = $_.Role
        Source                      = "Active"
        hasStrongMFA                = $_.hasStrongMFA
        AdminRiskScore              = $_.AdminRiskScore
        AdminRiskLevel              = $_.AdminRiskLevel
        lastSignInDateTime          = $_.lastSignInDateTime
        ProductivityServicesEnabled = $_.ProductivityServicesEnabled
    }
}

# -------------------------
# Eligible Roles
# -------------------------
$allAdmins += $eligible | ForEach-Object {
    [PSCustomObject]@{
        UserPrincipalName           = $_.UserPrincipalName
        ObjectId                    = $_.Id
        Role                        = $_.EligibleRole
        Source                      = "Eligible"
        hasStrongMFA                = $_.hasStrongMFA
        AdminRiskScore              = $_.AdminRiskScore
        AdminRiskLevel              = $_.AdminRiskLevel
        lastSignInDateTime          = $_.lastSignInDateTime
        ProductivityServicesEnabled = $_.ProductivityServicesEnabled
    }
}

# -------------------------
# Azure Roles
# -------------------------
$allAdmins += $azroles | Where-Object ObjectType -eq "User" | ForEach-Object {
    [PSCustomObject]@{
        UserPrincipalName           = $_.SignInName
        ObjectId                    = $_.ObjectId
        Role                        = $_.RoleDefinitionName
        Source                      = "Azure"
        hasStrongMFA                = $false
        AdminRiskScore              = $_.AdminRiskScore
        AdminRiskLevel              = $_.AdminRiskLevel
        lastSignInDateTime          = $null
        ProductivityServicesEnabled = $false
    }
}

# -------------------------
# CLEAN + DEDUPE (INCLUDED!)
# -------------------------
$topRiskAdmins = $allAdmins |
Where-Object { 
    $_.displayName -and 
    $_.displayName.Trim() -ne ""
} |
Sort-Object AdminRiskScore -Descending |
Select-Object -First 10 displayName, Role, AdminRiskScore, AdminRiskLevel
$topRiskAdmins = $topRiskAdmins | Sort-Object AdminRiskScore -Descending


$allAdmins = $allAdmins |
Where-Object { $_ -and $_.UserPrincipalName } |
Group-Object ObjectId |
ForEach-Object {
    $_.Group | Sort-Object AdminRiskScore -Descending | Select-Object -First 1
}


# ----------------------------------------------------
# Summary metrics
# ----------------------------------------------------
$totalAdmins = $allAdmins.Count

$criticalAdmins = ($allAdmins | Where-Object { $_.AdminRiskLevel -eq "Critical" }).Count
$highAdmins = ($allAdmins | Where-Object { $_.AdminRiskLevel -eq "High" }).Count
$mediumAdmins = ($allAdmins | Where-Object { $_.AdminRiskLevel -eq "Medium" }).Count
$lowAdmins = ($allAdmins | Where-Object { $_.AdminRiskLevel -eq "Low" }).Count

$noMFAAdmins = ($allAdmins | Where-Object { -not $_.hasStrongMFA }).Count

$inactiveAdmins = ($allAdmins | Where-Object {
        $_.lastSignInDateTime -and
        [datetime]$_.lastSignInDateTime -lt (Get-Date).AddDays(-90)
    }).Count

$productivityAdmins = ($allAdmins | Where-Object {
        $_.ProductivityServicesEnabled -eq $true
    }).Count


$allAdmins | Export-Excel `
    -Path $exportPath `
    -WorksheetName "All Admins" `
    -TableName AllAdminsTable `
    -AutoSize `



# MFA DATA
$mfaStatuses = @("Strong MFA", "No MFA")

$mfaTable = foreach ($status in $mfaStatuses) {

    if ($status -eq "Strong MFA") {
        $count = ($allAdmins | Where-Object { $_.hasStrongMFA }).Count
    }
    else {
        $count = ($allAdmins | Where-Object { -not $_.hasStrongMFA }).Count
    }

    [PSCustomObject]@{
        MFAStatus = $status
        Count     = $count
    }
}
$mfaTable | Export-Excel `
    -Path $exportPath `
    -WorksheetName "MFA Data" `
    -TableName 'MfaTable' `
    -AutoSize `
    -Append

# ----------------------------
# Risk Data
# -----------------------------

$riskLevels = @("Critical", "High", "Medium", "Low")

$riskTable = @(
    foreach ($level in $riskLevels) {

        $count = @($allAdmins | Where-Object { $_.AdminRiskLevel -eq $level }).Count

        [PSCustomObject]@{
            RiskLevel = $level
            Count     = $count
        }
    }
)

$riskTable | Format-Table
$riskTable | Get-Member
$riskTable.Count

$riskTable | Export-Excel `
    -Path $exportPath `
    -WorksheetName "Risk Data" `
    -TableName 'RiskTable' `
    -AutoSize `
    -Append

Write-Host "MFA rows: $($mfaTable.Count)"
Write-Host "Risk rows: $($riskTable.Count)"

# ---------------------------------
# TopRiskAdmins Data 
# ---------------------------------

$topRiskAdmins | Export-Excel `
    -Path $exportPath `
    -WorksheetName "TopRiskAdmin Data" `
    -TableName 'TopRiskAdmins' `
    -AutoSize `
    -Append






# ----------------------------------------------------
# Identity Posture Score
# ----------------------------------------------------

$total = $totalAdmins

# -----------------------------
# Calculate normalized score
# -----------------------------
$postureScore = 100

if ($total -gt 0) {
    $postureScore -= (($noMFAAdmins / $total) * 40)
    $postureScore -= (($criticalAdmins / $total) * 30)
    $postureScore -= (($inactiveAdmins / $total) * 20)
    $postureScore -= (($productivityAdmins / $total) * 10)
}

$postureScore = [math]::Round($postureScore, 0)
if ($postureScore -lt 0) { $postureScore = 0 }

# -----------------------------
# Level (human readable)
# -----------------------------
switch ($postureScore) {
    { $_ -ge 90 } { $postureLevel = "Excellent"; break }
    { $_ -ge 75 } { $postureLevel = "Good"; break }
    { $_ -ge 50 } { $postureLevel = "Moderate Risk"; break }
    default { $postureLevel = "High Risk" }
}

# -----------------------------
# Grade (executive friendly)
# -----------------------------
switch ($postureScore) {
    { $_ -ge 85 } { $grade = "A"; break }
    { $_ -ge 70 } { $grade = "B"; break }
    { $_ -ge 50 } { $grade = "C"; break }
    { $_ -ge 30 } { $grade = "D"; break }
    default { $grade = "F" }
}

# -----------------------------
# UX label 
# -----------------------------
$postureLabel = "$postureScore ($postureLevel - Grade $grade)"

# -----------------------------
# Export Object
# -----------------------------
$posture = @(
    [PSCustomObject]@{ Metric = "Privileged Identity Security Score"; Value = $postureScore }
    [PSCustomObject]@{ Metric = "Posture Level"; Value = $postureLevel }
    [PSCustomObject]@{ Metric = "Security Grade"; Value = $grade }
    [PSCustomObject]@{ Metric = "Display"; Value = $postureLabel }
)
# ----------------------------------------------------
# Build summary dataset
# ----------------------------------------------------

$summary = @(
    [PSCustomObject]@{ Metric = "Total Administrators"; Value = $totalAdmins }
    [PSCustomObject]@{ Metric = "Critical Risk Admins"; Value = $criticalAdmins }
    [PSCustomObject]@{ Metric = "High Risk Admins"; Value = $highAdmins }
    [PSCustomObject]@{ Metric = "Medium Risk Admins"; Value = $mediumAdmins }
    [PSCustomObject]@{ Metric = "Low Risk Admins"; Value = $lowAdmins }
    [PSCustomObject]@{ Metric = "Admins without Strong MFA"; Value = $noMFAAdmins }
    [PSCustomObject]@{ Metric = "Admins with Productivity Services"; Value = $productivityAdmins }
    [PSCustomObject]@{ Metric = "Inactive Admins (>90 days)"; Value = $inactiveAdmins }


)

# ----------------------------------------------------
# Dashboard dataset
# ----------------------------------------------------

$dashboard = @(
    [PSCustomObject]@{ Metric = "Total Admins"; Value = $totalAdmins }
    [PSCustomObject]@{ Metric = "Critical Admins"; Value = $criticalAdmins }
    [PSCustomObject]@{ Metric = "High Risk Admins"; Value = $highAdmins }
    [PSCustomObject]@{ Metric = "Medium Risk Admins"; Value = $mediumAdmins }
    [PSCustomObject]@{ Metric = "Low Risk Admins"; Value = $lowAdmins }
    [PSCustomObject]@{ Metric = "Admins without MFA"; Value = $noMFAAdmins }
    [PSCustomObject]@{ Metric = "Inactive Admins"; Value = $inactiveAdmins }
    [PSCustomObject]@{ Metric = "Admins with Productivity Services"; Value = $productivityAdmins }
)

$riskDistribution = @(
    [PSCustomObject]@{ RiskLevel = "Critical"; Count = $criticalAdmins }
    [PSCustomObject]@{ RiskLevel = "High"; Count = $highAdmins }
    [PSCustomObject]@{ RiskLevel = "Medium"; Count = $mediumAdmins }
    [PSCustomObject]@{ RiskLevel = "Low"; Count = $lowAdmins }
)
Write-Host "Exporting data to Excel..." -ForegroundColor Cyan




# Administrators
$administrators | Export-Excel `
    -NoNumberConversion * `
    -Path $exportPath `
    -WorksheetName "Administrators" `
    -TableName Administrators `
    -FreezeTopRow `
    -AutoSize `
    -TableStyle Medium2 `
    -Append

# Eligible Roles
$eligible | Export-Excel `
    -NoNumberConversion * `
    -Path $exportPath `
    -WorksheetName "Eligible Roles" `
    -TableName EligibleRoles `
    -FreezeTopRow `
    -AutoSize `
    -Append `
    -TableStyle Medium2

# Azure Roles
$azroles | Export-Excel `
    -NoNumberConversion * `
    -Path $exportPath `
    -WorksheetName "Azure Roles" `
    -TableName AzureRoles `
    -FreezeTopRow `
    -AutoSize `
    -Append `
    -TableStyle Medium2



# ----------------------------------------------------
# Dashboard Charts
# ----------------------------------------------------

$posture | Export-Excel `
    -Path $exportPath `
    -WorksheetName "Identity Posture" `
    -TableName IdentityPosture `
    -AutoSize `
    -Append `
    -TableStyle Medium2

<# $riskDistribution | Export-Excel `
    -Path $exportPath `
    -WorksheetName "Identity Dashboard" `
    -TableName RiskDistribution `
    -Append `
    -AutoSize #>





Write-Host "Adding conditional formatting..." -ForegroundColor Cyan

$excel = Open-ExcelPackage $exportPath
$adminRows = $administrators.Count + 1
$eligibleRows = $eligible.Count + 1
$azRows = $azroles.Count + 1
$topRows = $topRiskAdmins.Count + 1







# Administrators sheet
$ws = $excel.Workbook.Worksheets["Administrators"]

Add-ConditionalFormatting -Worksheet $ws -Address "F2:F$adminRows" `
    -RuleType ContainsText `
    -ConditionValue "Critical" `
    -BackgroundColor Red

Add-ConditionalFormatting -Worksheet $ws -Address "F2:F$adminRows" `
    -RuleType ContainsText `
    -ConditionValue "High" `
    -BackgroundColor Orange

Add-ConditionalFormatting -Worksheet $ws -Address "F2:F$adminRows" `
    -RuleType ContainsText `
    -ConditionValue "Medium" `
    -BackgroundColor Yellow

Add-ConditionalFormatting -Worksheet $ws -Address "F2:F$adminRows" `
    -RuleType ContainsText `
    -ConditionValue "Low" `
    -BackgroundColor LightGreen

# Eligible Sheet 

$ws = $excel.Workbook.Worksheets["Eligible Roles"]

Add-ConditionalFormatting -Worksheet $ws -Address "H2:H$eligibleRows" `
    -RuleType ContainsText `
    -ConditionValue "Critical" `
    -BackgroundColor Red

Add-ConditionalFormatting -Worksheet $ws -Address "H2:H$eligibleRows" `
    -RuleType ContainsText `
    -ConditionValue "High" `
    -BackgroundColor Orange

Add-ConditionalFormatting -Worksheet $ws -Address "H2:H$eligibleRows" `
    -RuleType ContainsText `
    -ConditionValue "Medium" `
    -BackgroundColor Yellow

Add-ConditionalFormatting -Worksheet $ws -Address "H2:H$eligibleRows" `
    -RuleType ContainsText `
    -ConditionValue "Low" `
    -BackgroundColor LightGreen
# Azure Roles  Sheet 

$ws = $excel.Workbook.Worksheets["Azure Roles"]

Add-ConditionalFormatting -Worksheet $ws -Address "I2:I$azRows" `
    -RuleType ContainsText `
    -ConditionValue "Critical" `
    -BackgroundColor Red

Add-ConditionalFormatting -Worksheet $ws -Address "I2:I$azRows" `
    -RuleType ContainsText `
    -ConditionValue "High" `
    -BackgroundColor Orange

Add-ConditionalFormatting -Worksheet $ws -Address "I2:I$azRows" `
    -RuleType ContainsText `
    -ConditionValue "Medium" `
    -BackgroundColor Yellow

Add-ConditionalFormatting -Worksheet $ws -Address "I2:I$azRows" `
    -RuleType ContainsText `
    -ConditionValue "Low" `
    -BackgroundColor LightGreen







# -----------------------------
# Traffic light icons for score
# -----------------------------

<#Add-ConditionalFormatting `
    -Worksheet $ws2 `
    -Address "D2:D$topRows" `
    -ThreeIconsSet TrafficLights1 `
    -Reverse #>
# ----------------------------------------------------
# BUILD COMPLETE SELF-CONTAINED DASHBOARD
# ----------------------------------------------------


$ws = $excel.Workbook.Worksheets["Identity Dashboard"]
if (-not $ws) {
    $ws = $excel.Workbook.Worksheets.Add("Identity Dashboard")
}

#$ws.Cells.Clear()
#$ws.Drawings.Clear()
#$ws.Tables.Clear()
$ws.View.ShowGridLines = $false

# ----------------------------------------------------
# TITLE
# ----------------------------------------------------
$ws.Cells["A1"].Value = "Entra Administrator Identity Security Dashboard"
$ws.Cells["A1:X1"].Merge = $true
$ws.Cells["A1"].Style.Font.Size = 28
$ws.Cells["A1"].Style.Font.Bold = $true
$ws.Cells["A1"].Style.HorizontalAlignment = "Center"

$ws.Cells["A1"].Style.Fill.PatternType = "Solid"
$ws.Cells["A1"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(122, 31, 162))
$ws.Cells["A1"].Style.Font.Color.SetColor([System.Drawing.Color]::White)

# ----------------------------------------------------
# KPI CARDS
# ----------------------------------------------------
function Set-KPI {
    param($range)

    $ws.Cells[$range].Style.Fill.PatternType = "Solid"
    $ws.Cells[$range].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(245, 245, 245))
    $ws.Cells[$range].Style.HorizontalAlignment = "Center"
    $ws.Cells[$range].Style.VerticalAlignment = "Center"
    $ws.Cells[$range].Style.Border.BorderAround("Medium")
}
$ws.Cells["B3:C3"].Merge = $true
$ws.Cells["B4:C4"].Merge = $true
$ws.Cells["B3"].Value = "Total Admins"
$ws.Cells["B4"].Value = $totalAdmins

$ws.Cells["E3:F3"].Merge = $true
$ws.Cells["E4:F4"].Merge = $true
$ws.Cells["E3"].Value = "Critical"
$ws.Cells["E4"].Value = $criticalAdmins

$ws.Cells["H3:I3"].Merge = $true
$ws.Cells["H4:I4"].Merge = $true
$ws.Cells["H3"].Value = "No MFA"
$ws.Cells["H4"].Value = $noMFAAdmins

$ws.Cells["K3:L3"].Merge = $true
$ws.Cells["K4:L4"].Merge = $true
$ws.Cells["K3"].Value = "Inactive"
$ws.Cells["K4"].Value = $inactiveAdmins

Set-KPI "B3:C4"
Set-KPI "E3:F4"
Set-KPI "H3:I4"
Set-KPI "K3:L4"

$ws.Cells["B3:L3"].Style.Font.Size = 10
$ws.Cells["B3:L3"].Style.Font.Bold = $true
$ws.Cells["B4:L4"].Style.Font.Size = 18
$ws.Cells["B4:L4"].Style.Font.Bold = $true

# ----------------------------------------------------
# POSTURE SCORE
# ----------------------------------------------------


# Reset
$ws.Cells["F5:J9"].Merge = $false

# merge
$ws.Cells["F5:J9"].Merge = $true

$ws.Cells["F5"].Value = "Privileged Identity Security Score`n`n$postureScore`n$postureLevel (Grade $grade)"

$ws.Cells["F5"].Style.WrapText = $true
$ws.Cells["F5"].Style.HorizontalAlignment = "Center"
$ws.Cells["F5"].Style.VerticalAlignment = "Center"

$ws.Cells["F5"].Style.Font.Size = 22
$ws.Cells["F5"].Style.Font.Bold = $true

$ws.Cells["F5:J9"].Style.Fill.PatternType = "Solid"
$ws.Cells["F5:J9"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(230, 240, 255))

$ws.Cells["F5:J9"].Style.Border.BorderAround("Thick")

# Color based on score
if ($postureScore -ge 75) { $color = [System.Drawing.Color]::Green }
elseif ($postureScore -ge 50) { $color = [System.Drawing.Color]::Orange }
else { $color = [System.Drawing.Color]::Red }

$ws.Cells["F5"].Style.Font.Color.SetColor($color)



# Style background
$ws.Cells["F6:H8"].Style.Fill.PatternType = "Solid"
$ws.Cells["F6:H8"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(245, 245, 245))

# Border
$ws.Cells["F6:H8"].Style.Border.BorderAround("Medium")

# Alignment
$ws.Cells["F6:H8"].Style.HorizontalAlignment = "Center"
$ws.Cells["F6:H8"].Style.VerticalAlignment = "Center"

# Font styles
$ws.Cells["F6"].Style.Font.Size = 10
$ws.Cells["F6"].Style.Font.Bold = $true

$ws.Cells["F7"].Style.Font.Size = 28
$ws.Cells["F7"].Style.Font.Bold = $true

$ws.Cells["F8"].Style.Font.Size = 10

if ($postureScore -ge 75) { $color = [System.Drawing.Color]::Green }
elseif ($postureScore -ge 50) { $color = [System.Drawing.Color]::Orange }
else { $color = [System.Drawing.Color]::Red }

$ws.Cells["F7"].Style.Font.Color.SetColor($color)





# ----------------------------------------------------
# CHARTS
# ----------------------------------------------------

# Risk
Add-ExcelChart `
    -Worksheet $ws `
    -ChartType Doughnut `
    -Title "Admin Risk Distribution" `
    -XRange "RiskTable[RiskLevel]" `
    -YRange "RiskTable[Count]" `
    -Row 10 `
    -Column 1 `
    -Width 450 `
    -Height 300

# MFA
Add-ExcelChart `
    -Worksheet $ws `
    -ChartType Doughnut `
    -Title "MFA Coverage" `
    -XRange "MfaTable[MFAStatus]" `
    -YRange "MfaTable[Count]" `
    -Row 10 `
    -Column 9 `
    -Width 450 `
    -Height 300

# Top admins
Add-ExcelChart `
    -Worksheet $ws `
    -ChartType BarClustered `
    -Title "Top 10 Highest Risk Admins" `
    -XRange "TopRiskAdmins[displayName]" `
    -YRange "TopRiskAdmins[AdminRiskScore]" `
    -Row 28 `
    -Column 1 `
    -Width 1400 `
    -Height 450

Close-ExcelPackage -ExcelPackage  $excel

Write-Host "✅ Export completed successfully: $exportPath" -ForegroundColor Green
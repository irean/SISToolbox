

function Test-Module {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [String]$Name
    )

    Write-Host "Checking module '$Name'..." -ForegroundColor Cyan
    if (-not (Get-Module $Name)) {
        Write-Host "Module '$Name' not imported, attempting import..." -ForegroundColor Yellow
        try {
            if ($Name -eq 'Microsoft.Graph') {
                Write-Host "Importing Microsoft.Graph (this may take a while)..."
            }
            Import-Module $Name -ErrorAction Stop
        }
        catch {
            Write-Host "❌ Module '$Name' not found. Installing..." -ForegroundColor Red
            Install-Module $Name -Scope CurrentUser -AllowClobber -Force -AcceptLicense -SkipPublisherCheck
            Write-Host "Importing module '$Name' after install..." -ForegroundColor Cyan
            Import-Module $Name -ErrorAction Stop 
        }
    } 
    else {
        Write-Host "✅ Module '$Name' is already imported." -ForegroundColor Green
    }   
    <#
.SYNOPSIS
    Verifies and imports required PowerShell modules.

.DESCRIPTION
    This function checks whether a specified module is imported.
    If not, it attempts to import or install it as needed.

.PARAMETER Name
    The name of the module to verify.

.EXAMPLE
    Test-Module -Name Microsoft.Graph

.NOTES
    Required for most Microsoft Graph operations.
#>
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


function Get-AccessPackageAssignmentsTargets {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [String]$AccessPackageId
    )

    $uri = "https://graph.microsoft.com/v1.0/identityGovernance/entitlementManagement/assignments?`$expand=target,accessPackage,assignmentPolicy&`$filter=accessPackage/id eq '$AccessPackageId' and state eq 'Delivered'"
    $assignments = igall $uri

    $assignments | ForEach-Object {
        [PSCustomObject]@{
            AccessPackageName         = $_.accessPackage.displayName
            AccessPackagePolicyName   = $_.assignmentPolicy.displayName
            TargetObjectId            = $_.target.id
            PrincipalName             = $_.target.userPrincipalName
            DisplayName               = $_.target.displayName
            State                     = $_.state
            ScheduleStart             = $_.schedule.startDateTime
            Status                    = $_.status
            AccessPackageAssignmentId = $_.id
        }
    }
    <#
.SYNOPSIS
    Retrieves all delivered assignments and their target users for a specific Access Package.

.DESCRIPTION
    The Get-AccessPackageAssignmentsTargets function queries Microsoft Graph's 
    Entitlement Management API to retrieve assignment data for a given Access Package.
    It filters for assignments in the 'Delivered' state and expands related properties, 
    including target user details, access package, and assignment policy information.

    The output includes key user and assignment details such as:
      - Access Package name
      - Assignment policy name
      - Target object ID
      - User principal name
      - Display name
      - Assignment state and status
      - Assignment start schedule
      - Access Package Assignment ID

.PARAMETER AccessPackageId
    The Object ID of the Access Package to query.
    This is required to filter the Graph API request and return the relevant assignments.

.EXAMPLE
    PS> Get-AccessPackageAssignmentsTargets -AccessPackageId "b3a77f84-6a3d-44b1-9f50-d32c17346a31"

    AccessPackageName    : Finance Onboarding
    AccessPackagePolicy  : Default Policy
    TargetObjectId       : 8c5ab4e3-14a2-47f8-a6a3-9ef8c97d3a90
    PrincipalName        : jane.doe@company.com
    DisplayName          : Jane Doe
    State                : Delivered
    ScheduleStart        : 2025-01-10T07:00:00Z
    Status               : Delivered
    AccessPackageAssignmentId : 1f441bb1-1d6e-466b-a73a-3bdf985f8ea8

    This example retrieves all current user assignments for the specified Access Package.

.EXAMPLE
    PS> $assignments = Get-AccessPackageAssignmentsTargets -AccessPackageId $packageId
    PS> $assignments | Where-Object { $_.PrincipalName -like '*@company.com' }

    Retrieves all assignments for a specific Access Package and filters results to users 
    within the domain.

.INPUTS
    System.String

.OUTPUTS
    System.Management.Automation.PSCustomObject
        Each object represents a user’s assignment to an Access Package.

.REQUIRED_SCOPES
    EntitlementManagement.Read.All
    User.Read.All

.NOTES
    Author: Sandra Saluti
    Version: 1.0
    Tags: Microsoft Graph, Entitlement Management, Governance, Access Packages
    Date: 2025-11-10

    This function uses the helper function `igall` to perform Graph API pagination and 
    retrieve all results.

.LINK
    https://learn.microsoft.com/en-us/graph/api/resources/entitlementmanagement-root

#>

}



function Compare-UsersToAccessPackageAssignments {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] [String]$AccessPackageId,
        [Parameter(Mandatory = $true)] [Array]$UserList
    )

    $assignedUsers = Get-AccessPackageAssignmentsTargets -AccessPackageId $AccessPackageId

    $UserList | ForEach-Object {
        $match = $assignedUsers | Where-Object {
            $_.TargetObjectId -eq $_.ObjectId -or $_.PrincipalName -eq $_.PrincipalName
        }

        if ($match) {
            $match | ForEach-Object {
                [PSCustomObject]@{
                    User                      = $_.PrincipalName
                    Assigned                  = $true
                    AccessPackageName         = $_.AccessPackageName
                    AccessPackagePolicyName   = $_.AccessPackagePolicyName
                    TargetObjectId            = $_.TargetObjectId
                    DisplayName               = $_.DisplayName
                    EmployeeId                = $_.EmployeeId
                    EmployeeHireDate          = $_.EmployeeHireDate
                    JobTitle                  = $_.JobTitle
                    AccountEnabled            = $_.AccountEnabled
                    State                     = $_.State
                    ScheduleStart             = $_.ScheduleStart
                    Status                    = $_.Status
                    AccessPackageAssignmentId = $_.AccessPackageAssignmentId
                }
            }

        }
        else {
            [PSCustomObject]@{
                User                    = $_.PrincipalName
                Assigned                = $false
                AccessPackageName       = $null
                AccessPackagePolicyName = $null
                TargetObjectId          = $null
                DisplayName             = $_.DisplayName
                EmployeeId              = $_.EmployeeId
                EmployeeHireDate        = $_.EmployeeHireDate
                JobTitle                = $_.JobTitle
                AccountEnabled          = $_.AccountEnabled
            }
        }
    }
    <#
.SYNOPSIS
    Compares a list of users to those assigned to a specific Access Package.

.DESCRIPTION
    The Compare-UsersToAccessPackageAssignments function checks whether each user in 
    the provided list is currently assigned to a specified Access Package. 

    It uses the Get-AccessPackageAssignmentsTargets function to retrieve all existing 
    assignments for the given Access Package, then compares them against the provided 
    user list by ObjectId or UserPrincipalName.

    The result is a collection of objects indicating whether each user is assigned or not,
    along with detailed assignment and user information.

.PARAMETER AccessPackageId
    The Object ID of the Access Package to compare users against.

.PARAMETER UserList
    An array of user objects to compare. 
    Each object should contain at least:
      - ObjectId
      - PrincipalName
      - DisplayName
      - EmployeeId
      - EmployeeHireDate
      - JobTitle
      - AccountEnabled

.EXAMPLE
    Compare-UsersToAccessPackageAssignments -AccessPackageId "b3a77f84-6a3d-44b1-9f50-d32c17346a31" -UserList $users

    This example compares all Danish users in Entra ID with the list of users assigned to 
    a specific Access Package and returns assignment status for each.

.EXAMPLE
    PS> $report = Compare-UsersToAccessPackageAssignments -AccessPackageId $packageId -UserList $onboardedUsers
    PS> $report | Where-Object { -not $_.Assigned }

    Retrieves all users not currently assigned to the specified Access Package.

.INPUTS
    System.String
    System.Array

.OUTPUTS
    System.Management.Automation.PSCustomObject
        Each object includes assignment state and key user details.

.REQUIRED_SCOPES
    EntitlementManagement.Read.All
    User.Read.All

.NOTES
    Author: Sandra Saluti
    Version: 1.0
    Tags: Microsoft Graph, Entitlement Management, Governance, Access Packages, Comparison
    Date: 2025-11-10

    This function depends on Get-AccessPackageAssignmentsTargets to fetch the 
    current assignments from Microsoft Graph.

.LINK
    Get-AccessPackageAssignmentsTargets
    https://learn.microsoft.com/en-us/graph/api/resources/entitlementmanagement-root
    Get-Help about_Functions
#>

}



function Get-UsersDynamic {
    [CmdletBinding(DefaultParameterSetName = "Filter")]
    param (
        [Parameter(ParameterSetName = "Filter")] [string]$EmployeeIdStartsWith,
        [Parameter(ParameterSetName = "Filter")] [string]$Country,
        [Parameter(ParameterSetName = "Filter")] [string]$Department,
        [Parameter(ParameterSetName = "Filter")] [string]$CompanyName,
        [Parameter(ParameterSetName = "Filter")] [switch]$EmployeeIdNotNull,
        [Parameter(ParameterSetName = "Filter")] [switch]$EmployeeLeaveDateTimeNotNull,
        [Parameter(ParameterSetName = "Group")]  [string]$GroupId,
        [Parameter(ParameterSetName = "Manual")] [array]$Users
    )

    switch ($PSCmdlet.ParameterSetName) {
        "Filter" {
            $filterParts = @()
            if ($EmployeeIdStartsWith) {
                $filterParts += "startsWith(employeeId, '$EmployeeIdStartsWith')" 
            }
            if ($Country) {
                $filterParts += "country eq '$Country'" 
            }
            if ($Department) {
                $filterParts += "department eq '$Department'"
            }
            if ($companyName) {
                $filterParts += "companyName eq '$companyName'"
            }
            $filterParts += "accountEnabled eq true"

            $filterQuery = [string]::Join(' and ', $filterParts)
            $uri = "https://graph.microsoft.com/v1.0/users?`$filter=$filterQuery&`$select=id,displayName,userPrincipalName,employeeId,employeeHireDate,employeeLeaveDateTime,jobTitle,accountEnabled,department,companyName&`$count=true"
            
            Write-Host "🔎 Querying users with: $filterQuery" -ForegroundColor Yellow
            $data = igall $uri -Eventual

            if ($EmployeeIdNotNull) { $data = $data | Where-Object { $_.employeeId } }
            if ($EmployeeLeaveDateTimeNotNull) { $data = $data | Where-Object { $_.employeeLeaveDateTime } }
        }

        "Group" {
            Write-Host "📂 Getting users from group ID: $GroupId" -ForegroundColor Yellow
            $uri = "https://graph.microsoft.com/v1.0/groups/$GroupId/members/microsoft.graph.user?`$select=id,displayName,userPrincipalName,employeeId,employeeHireDate,employeeLeaveDateTime,jobTitle,accountEnabled,department,companyName&`$count=true"
            $data = igall $uri -Eventual
        }

        "Manual" {
            Write-Host "✋ Using manually provided user list" -ForegroundColor Yellow
            $data = $Users
        }
    }

    return $data | ForEach-Object {
        [PSCustomObject]@{
            ObjectId              = $_.id
            DisplayName           = $_.displayName
            PrincipalName         = $_.userPrincipalName
            EmployeeId            = $_.employeeId
            EmployeeHireDate      = $_.employeeHireDate
            EmployeeLeaveDateTime = $_.employeeLeaveDateTime
            JobTitle              = $_.jobTitle
            AccountEnabled        = $_.accountEnabled
            Department            = $_.Department
            CompanyName           = $_.CompanyName
        }
    }
    <#
.SYNOPSIS
    Retrieves Entra ID users dynamically based on filters, group membership, or manual input.

.DESCRIPTION
    The Get-UsersDynamic function provides a flexible way to retrieve user objects from Microsoft Entra ID (Azure AD).  
    It supports three parameter sets:
      - **Filter**: Build an OData filter dynamically (e.g., by country, department, company name, or employeeId prefix).
      - **Group**: Retrieve users directly from a Microsoft 365 or Entra ID group.
      - **Manual**: Accept a pre-provided array of user objects.

    The output is standardized into consistent PSCustomObjects with key attributes like DisplayName, JobTitle, CompanyName, and more.


.PARAMETER Country
    Filters users by the `country` attribute.  
    Example: `-Country "Denmark"`.

.PARAMETER Department
    Filters users by the `department` attribute.

.PARAMETER CompanyName
    Filters users by the `companyName` attribute.

.PARAMETER EmployeeIdNotNull
    When specified, removes users without an EmployeeId from the results after Graph retrieval.

.PARAMETER EmployeeLeaveDateTimeNotNull
    When specified, removes users without an EmployeeLeaveDateTime from the results after Graph retrieval.

.PARAMETER GroupId
    Specifies a group ObjectId to retrieve members from.  
    Used with the `-Group` parameter set.

.PARAMETER Users
    Manually provides a list of user objects for comparison or testing.  
    Used with the `-Manual` parameter set.

.EXAMPLE
    PS> Get-UsersDynamic -Country "Denmark" -CompanyName "Company ApS"

    Retrieves all enabled users in Denmark working for "Company ApS".


.EXAMPLE
    PS> Get-UsersDynamic -GroupId "f8b2b7e0-2b9e-4f32-9a35-5ff7f68ac12b"

    Retrieves all members of the specified group.



.INPUTS
    System.String
    System.Array

.OUTPUTS
    System.Management.Automation.PSCustomObject
        Each object includes:
          - ObjectId
          - DisplayName
          - PrincipalName
          - EmployeeId
          - EmployeeHireDate
          - EmployeeLeaveDateTime
          - JobTitle
          - AccountEnabled
          - Department
          - CompanyName

.REQUIRED_SCOPES
    User.Read.All
    Directory.Read.All

.NOTES
    Author: Sandra Saluti
    Version: 1.0
    Date: 2025-11-10
    Tags: Microsoft Graph, Entra ID, Azure AD, User Management, Dynamic Query

    This function relies on a helper function `igall` to handle paginated Graph API calls and eventual consistency.
    Only enabled accounts (`accountEnabled eq true`) are retrieved by default.

.LINK
    https://learn.microsoft.com/en-us/graph/api/resources/user
    Get-Help about_Functions
#>


}



function Select-FolderPath {
    [CmdletBinding()]
    param()

    Write-Host "--------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host "📂 Please select a folder where the report will be saved." -ForegroundColor Cyan
    Write-Host "⚠️  The folder selection window may appear behind other open windows." -ForegroundColor Yellow
    Write-Host "If you don't see it, try minimizing other windows." -ForegroundColor Yellow
    Write-Host "--------------------------------------------------------" -ForegroundColor DarkGray

    Add-Type -AssemblyName System.Windows.Forms

    $FileBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{ 
        Description         = "Select a folder for the report export"
        RootFolder          = [Environment+SpecialFolder]::Desktop
        ShowNewFolderButton = $true
    }

    $form = New-Object System.Windows.Forms.Form -Property @{ TopMost = $true }
    $result = $FileBrowser.ShowDialog($form)

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $folder = $FileBrowser.SelectedPath
        Write-Host "✅ Export folder selected: $folder" -ForegroundColor Green
        return $folder
    }
    else {
        Write-Host "❌ No folder selected. Exiting script." -ForegroundColor Red
        return $null
    }
    <#
.SYNOPSIS
    Opens a folder picker dialog for selecting an export folder.

.DESCRIPTION
    Displays a Windows folder selection dialog and returns the chosen path.
    The dialog is forced to the top of the screen to prevent it from opening behind other windows.

.EXAMPLE
    $folderPath = Select-FolderPath

.RETURNS
    [string] The selected folder path, or $null if cancelled.
#>
}


function Compare-UsersToAccessPackageAssignmentsWithProgress {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$AccessPackageId,

        [Parameter(Mandatory = $true)]
        [array]$UserList
    )

    Write-Host "`n Starting user comparison against Access Package assignments..." -ForegroundColor Yellow
    Write-Progress -Activity "🔎 Fetching Access Package assignments from Microsoft Graph" -Status "Please wait..." -PercentComplete 10

    # --- Step 1: Retrieve assignments ---
    $uri = "https://graph.microsoft.com/v1.0/identityGovernance/entitlementManagement/assignments?`$expand=target,accessPackage,assignmentPolicy&`$filter=accessPackage/id eq '$AccessPackageId' and state eq 'Delivered'"

    Write-Host "Querying Microsoft Graph for assignments..." -ForegroundColor Cyan
    Write-Host "    → $uri" -ForegroundColor DarkGray

    try {
        $assignments = igall $uri
        Write-Host "✅ Retrieved $($assignments.Count) Access Package assignments." -ForegroundColor Green
    }
    catch {
        Write-Host "❌ Failed to retrieve Access Package assignments. Check permissions or package ID." -ForegroundColor Red
        throw
    }

    # --- Step 2: Begin user comparison ---
    Write-Progress -Activity "Preparing comparison" -Status "Initializing user comparison..." -PercentComplete 50
    Start-Sleep -Milliseconds 500

    $totalUsers = $UserList.Count
    $compareResults = @()

    Write-Host "👥 Beginning comparison for $totalUsers users..." -ForegroundColor Cyan
    Start-Sleep -Milliseconds 300

    # --- Step 3: Compare users ---
    for ($i = 0; $i -lt $totalUsers; $i++) {
        $user = $UserList[$i]
        $progressPercent = [math]::Round((($i + 1) / $totalUsers) * 100, 2)

        Write-Progress -Activity "Comparing users to Access Package assignments" `
            -Status "Processing $($user.DisplayName) ($($i + 1)/$totalUsers)" `
            -PercentComplete $progressPercent

        # Match by ObjectId or UPN
        $match = $assignments | Where-Object {
            $_.target.id -eq $user.ObjectId -or $_.target.userPrincipalName -eq $user.PrincipalName
        }

        if ($match) {
            foreach ($m in $match) {
                $compareResults += [pscustomobject]@{
                    User                      = $user.PrincipalName
                    Assigned                  = $true
                    AccessPackageName         = $m.accessPackage.displayName
                    AccessPackagePolicyName   = $m.assignmentPolicy.displayName
                    TargetObjectId            = $m.target.id
                    DisplayName               = $user.DisplayName
                    EmployeeId                = $user.EmployeeId
                    EmployeeHireDate          = $user.EmployeeHireDate
                    EmployeeLeaveDateTime     = $user.EmployeeLeaveDateTime
                    JobTitle                  = $user.JobTitle
                    AccountEnabled            = $user.AccountEnabled
                    Department                = $user.Department
                    CompanyName               = $user.CompanyName
                    State                     = $m.state
                    ScheduleStart             = $m.schedule.startDateTime
                    Status                    = $m.status
                    AccessPackageAssignmentId = $m.id
                }
            }
        }
        else {
            $compareResults += [pscustomobject]@{
                User                      = $user.PrincipalName
                Assigned                  = $false
                AccessPackageName         = $null
                AccessPackagePolicyName   = $null
                TargetObjectId            = $null
                DisplayName               = $user.DisplayName
                EmployeeId                = $user.EmployeeId
                EmployeeHireDate          = $user.EmployeeHireDate
                EmployeeLeaveDateTime     = $user.EmployeeLeaveDateTime
                JobTitle                  = $user.JobTitle
                AccountEnabled            = $user.AccountEnabled
                Department                = $user.Department
                CompanyName               = $user.CompanyName
                State                     = $null
                ScheduleStart             = $null
                Status                    = $null
                AccessPackageAssignmentId = $null
            }
        }
    }

    Write-Progress -Activity "Comparing users to Access Package assignments" -Completed
    Write-Host "✅ Comparison complete for $totalUsers users." -ForegroundColor Green

    # --- Step 4: Summary ---
    $assignedCount = ($compareResults | Where-Object { $_.Assigned }).Count
    $notAssignedCount = $totalUsers - $assignedCount

    Write-Host "Summary:" -ForegroundColor Cyan
    Write-Host "   → Assigned: $assignedCount" -ForegroundColor Green
    Write-Host "   → Not assigned: $notAssignedCount" -ForegroundColor Yellow

    return $compareResults

    <#
.SYNOPSIS
    Compares a list of users to their Access Package assignments in Microsoft Entra ID with progress feedback.

.DESCRIPTION
    The Compare-UsersToAccessPackageAssignmentsWithProgress function retrieves Access Package assignments
    from Microsoft Graph for a given Access Package ID and compares them to a provided user list.  
    It shows real-time progress during the comparison, allowing you to monitor the status as users are processed.  

    The output identifies which users are assigned (state = Delivered) and which are not, returning a detailed
    list of comparison results for further reporting or export.

.PARAMETER AccessPackageId
    The ObjectId of the Access Package to compare users against.

.PARAMETER UserList
    An array of user objects (from Entra ID or local data) to check for assignments.  
    Each object should contain at least `ObjectId` and `PrincipalName`.

.EXAMPLE
    PS> $users = Get-UsersDynamic -Country "Denmark" -CompanyName "Company ApS"
    PS> Compare-UsersToAccessPackageAssignmentsWithProgress -AccessPackageId "bcd12345-abcd-6789-ef00-1234567890ab" -UserList $users

    Compares all users from Denmark in company "Company ApS" against a specific Access Package
    and returns assignment results with live progress updates.

.EXAMPLE
    PS> $compareResults = Compare-UsersToAccessPackageAssignmentsWithProgress -AccessPackageId "abcd1234" -UserList $UserList
    PS> $compareResults | Where-Object { -not $_.Assigned }

    Retrieves users who are NOT assigned to the specified Access Package.

.OUTPUTS
    System.Management.Automation.PSCustomObject
        Each object includes:
          - User
          - Assigned
          - AccessPackageName
          - AccessPackagePolicyName
          - TargetObjectId
          - DisplayName
          - EmployeeId
          - EmployeeHireDate
          - EmployeeLeaveDateTime
          - JobTitle
          - AccountEnabled
          - Department
          - CompanyName
          - State
          - ScheduleStart
          - Status
          - AccessPackageAssignmentId

.REQUIRED_SCOPES
    EntitlementManagement.Read.All
    User.Read.All

.NOTES
    Author: Sandra Saluti
    Version: 1.0
    Date: 2025-11-10
    Tags: Entra ID, Entitlement Management, Graph API, Governance

    This function uses a helper function `igall` to handle paginated Graph API responses.
    It includes detailed Write-Progress and Write-Host output for visibility in longer comparisons.

.LINK
    https://learn.microsoft.com/en-us/graph/api/resources/accesspackageassignment
    https://learn.microsoft.com/en-us/graph/api/resources/entitlementmanagement-overview
#>
}


function Start-UserAccessPackageAudit {
    [CmdletBinding()]
    param()

    Write-Host "==========================================" -ForegroundColor DarkGray
    Write-Host "🔎  USER & ACCESS PACKAGE AUDIT STARTED  " -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor DarkGray

    # ---  Ensure required modules ---
    Test-Module -Name Microsoft.Graph.Authentication
    Test-Module -Name ImportExcel

    # --- Connect to Microsoft Graph with correct scopes ---
    $requiredScopes = @(
        "User.Read.All",
        "Organization.Read.All",
        "Group.Read.All",
        "EntitlementManagement.Read.All"
    )

    Write-Host "`n Connecting to Microsoft Graph..." -ForegroundColor Yellow
    Connect-MgGraph -Scopes $requiredScopes | Out-Null
    Write-Host "✅ Connected to Graph successfully." -ForegroundColor Green

    # ---  Get organization name for export naming ---
    Write-Host "`nFetching organization name..." -ForegroundColor Yellow
    $orgDisplayName = (Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/organization").value.displayName
    Write-Host "🏢 Organization: $orgDisplayName" -ForegroundColor Green

    # ---  Choose user retrieval method ---
    Write-Host "`nHow would you like to get users?" -ForegroundColor Yellow
    Write-Host "1️⃣  Filter (Country, EmployeeId, etc.)"
    Write-Host "2️⃣  Group (Members of specific group)"
    Write-Host "3️⃣  Manual (CSV import)"
    $choice = Read-Host "Enter choice (1-3)"

    switch ($choice) {
        1 {
            Write-Host "`n--- 🔍 FILTER PARAMETERS ---" -ForegroundColor Cyan
            $country = Read-Host "Enter country (or leave blank)"
            $department = Read-Host "Filter by department (or leave blank)"
            $companyName = Read-Host "Filter by companyName (or leave blank)"
            $useEmpIdNotNull = Read-Host "Filter EmployeeId not null? (y/n)"
            $useLeaveDateNotNull = Read-Host "Filter EmployeeLeaveDateTime not null? (y/n)"

            $params = @{}
            if ($country) {
                $params.Country = $country 
            }
            if ($employeePrefix) { 
                $params.EmployeeIdStartsWith = $employeePrefix 
            }
            if ($department) { 
                $params.Department = $department
            }
            if ($companyName) { 
                $params.companyName = $companyName
            }
            if ($useEmpIdNotNull -eq 'y') { 
                $params.EmployeeIdNotNull = $true 
            }
            if ($useLeaveDateNotNull -eq 'y') {
                $params.EmployeeLeaveDateTimeNotNull = $true 
            }

            $users = Get-UsersDynamic @params
        }

        2 {
            Write-Host "`n--- 👥 GROUP MEMBERS ---" -ForegroundColor Cyan
            $groupId = Read-Host "Enter Group ObjectId"
            $users = Get-UsersDynamic -GroupId $groupId
        }

        3 {
            Write-Host "`n--- 📝 MANUAL USER IMPORT ---" -ForegroundColor Cyan
            $path = Read-Host "Enter CSV path (or leave blank to cancel)"
            if ($path) {
                $users = Import-Csv $path
            }
            else {
                Write-Warning "No file provided. Exiting."
                return
            }
        }

        default {
            Write-Warning "Invalid selection. Exiting script."
            return
        }
    }

    if (-not $users) {
        Write-Warning "No users retrieved. Exiting script."
        return
    }

    Write-Host "`n✅ Retrieved $($users.Count) users for comparison." -ForegroundColor Green

    # --- Select Access Package ---
    Write-Host "`n--- 🎁 ACCESS PACKAGE SELECTION ---" -ForegroundColor Cyan
    $accessPackageId = Read-Host "Enter Access Package ObjectId"
    $accessPackage = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/identityGovernance/entitlementManagement/accessPackages/$accessPackageId"
    $accessPackageName = $accessPackage.displayName -replace '[^\w\-]', '_'
    Write-Host "📦 Selected Access Package: $($accessPackage.displayName)" -ForegroundColor Green

    # --- Compare users to Access Package assignments ---
    Write-Host "`n🔍 Comparing users to Access Package assignments..." -ForegroundColor Yellow
    $compareResults = Compare-UsersToAccessPackageAssignmentsWithProgress -UserList $users -AccessPackage $accessPackageid


    # --- Build summary and totals ---
    $summary = $compareResults |
    Group-Object AccessPackagePolicyName, Assigned |
    Select-Object @{n = 'AccessPackagePolicyName'; e = { $_.Group[0].AccessPackagePolicyName } },
    @{n = 'Assigned'; e = { $_.Group[0].Assigned } },
    @{n = 'Count'; e = { $_.Count } } |
    Sort-Object AccessPackagePolicyName

    $totalSummary = $compareResults |
    Group-Object Assigned |
    Select-Object @{n = 'Assigned'; e = { $_.Name } }, @{n = 'Count'; e = { $_.Count } }

    # --- Select export folder ---
    Write-Host "`n📁 Select output folder for Excel export..." -ForegroundColor Yellow
    $folderPath = Select-FolderPath
    if (-not $folderPath) {
        Write-Warning "No folder selected. Exiting script."
        return
    }

    $date = Get-Date -Format 'yyyy-MM-dd'
    $fileName = "$($orgDisplayName)-$($accessPackageName)-$date.xlsx" -replace '\s+', '_'
    $exportPath = Join-Path $folderPath $fileName

    # --- Export to Excel (helper function) ---
    Export-AccessPackageReportToExcel -CompareResults $compareResults `
        -Summary $summary -TotalSummary $totalSummary -ExportPath $exportPath

    Write-Host "==========================================" -ForegroundColor DarkGray
    Write-Host "✅  USER & ACCESS PACKAGE AUDIT COMPLETE  " -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor DarkGray
    <#
.SYNOPSIS
    Starts an audit comparing Entra ID users to their Access Package assignments.

.DESCRIPTION
    The Start-UserAccessPackageAudit function connects to Microsoft Graph and retrieves users based on one of three methods:
    1. Filtered search (e.g., by Country, Department, CompanyName, EmployeeId, etc.)
    2. Group membership
    3. Manual CSV import

    It then compares these users against a selected Access Package to determine assignment status,
    and exports a report to an Excel file with detailed comparison results and summaries.

.PARAMETER None
    This function takes no parameters. It runs interactively and prompts for input values such as:
    - Filter conditions (country, department, companyName, etc.)
    - Group ObjectId (if group mode is selected)
    - CSV path (if manual import mode is selected)
    - Access Package ObjectId
    - Export folder path

.EXAMPLE
    PS C:\> Start-UserAccessPackageAudit
    Launches the interactive audit wizard where you can select user retrieval method, filter options,
    access package, and output location.

.INPUTS
    None. User input is gathered via Read-Host prompts.

.OUTPUTS
    - Excel file (.xlsx) containing:
        • User comparison list (Assigned / Not Assigned)
        • Summary per Access Package policy
        • Total counts

.REQUIREMENTS
    - Microsoft Graph PowerShell SDK (`Microsoft.Graph.Authentication`)
    - ImportExcel module
    - Sufficient Graph API permissions:
        • User.Read.All
        • Organization.Read.All
        • Group.Read.All
        • EntitlementManagement.Read.All

.NOTES
    Author: Sandra Saluti
    Version: 1.0
    Updated: November 2025
   
#>

}

function Export-AccessPackageReportToExcel {
    param(
        [Parameter(Mandatory)][PSObject[]]$CompareResults,
        [Parameter(Mandatory)][PSObject[]]$Summary,
        [Parameter(Mandatory)][PSObject[]]$TotalSummary,
        [Parameter(Mandatory)][string]$ExportPath
    )

    # --- Export DetailedReport --
    $pkg = $CompareResults | Export-Excel -Path $ExportPath `
        -WorksheetName 'DetailedReport' `
        -TableStyle Medium2 -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow `
        -PassThru  `
        -TableName 'DetailedTable'


 
    # --- Export Summary sheet ---
    $Summary | Export-Excel -ExcelPackage $pkg `
        -WorksheetName 'Summary' `
        -AutoSize -AutoFilter `
        -TableStyle Medium1 `
        -TableName 'AccessPackageSummary' `
        -PassThru
    $rowcount = $summary.Count + 3


    $pivot = Add-PivotTable -ExcelPackage $pkg -Address $pkg.Summary.Cells["A$rowcount"] -PivotTableName "DetailedPivot" `
        -SourceWorksheet $pkg.DetailedReport -SourceRange $pkg.DetailedReport.Tables["DetailedTable"] `
        -PivotData @{"Assigned" = "count" } -PivotRows "Assigned" -IncludePivotChart -ChartType Pie3D -PassThru
    $pivot.GridDropZones = $false
    $pivot.DataFields[0].Name = "Count of Assigned"

    # --- Export UserNotAssigned --- 

    $pkg = $CompareResults | Where-Object {
        $_.Assigned -like 'FALSE'
    } |  Export-Excel -ExcelPackage $pkg  `
        -WorksheetName 'UserNotAssigned' `
        -TableStyle Medium2 -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow `
        -PassThru  `
        -TableName 'UserNotAssignedTable'
    Close-ExcelPackage -ExcelPackage $pkg -Show




    # --- Final Save --- 
    try {

        Write-Host "✅ Excel report created successfully: $ExportPath" -ForegroundColor Green
        Write-Host "Sheets included: Summary, DetailedReport" -ForegroundColor Cyan

    }
    catch {
        Write-Host "❌ Export failed — ExcelPackage not valid or save error." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Yellow
    }

    <#
.SYNOPSIS
    Exports Access Package audit results to an Excel workbook.

.DESCRIPTION
    The Export-AccessPackageReportToExcel function takes comparison results from an Access Package audit 
    (including detailed user data, summary statistics, and total counts) and exports them to an Excel file. 

    The exported workbook includes multiple worksheets:
    - **DetailedReport**: All users compared, with full attributes and assignment status.
    - **Summary**: Aggregated counts of users assigned vs not assigned per policy.
    - **UserNotAssigned**: Filtered list of users not assigned to the Access Package.
    - A **Pivot Table** and **3D Pie Chart** are also created in the Summary sheet for quick visualization.

.PARAMETER CompareResults
    The detailed comparison data returned from Compare-UsersToAccessPackageAssignmentsWithProgress.
    Each record represents a user and their assignment status.

.PARAMETER Summary
    Aggregated data grouped by AccessPackagePolicyName and assignment status.
    Used to populate the “Summary” worksheet.

.PARAMETER TotalSummary
    High-level totals grouped by assignment state (e.g., Assigned / Not Assigned).

.PARAMETER ExportPath
    The full file path (including filename) for the exported Excel workbook.
    Example: "C:\Reports\AccessPackageAudit.xlsx"

.EXAMPLE
    PS C:\> Export-AccessPackageReportToExcel `
        -CompareResults $compareResults `
        -Summary $summary `
        -TotalSummary $totalSummary `
        -ExportPath "C:\Reports\AccessPackageAudit.xlsx"

    Exports all results from the current audit to the specified Excel file. 
    The file will open automatically when the export completes.

.OUTPUTS
    Excel (.xlsx) file containing:
    - DetailedReport
    - Summary
    - UserNotAssigned
    - Pivot table & chart

.REQUIREMENTS
    - ImportExcel module
    - Valid Excel file path for export

.NOTES
    Author: Sandra Saluti
    Version: 1.0
    Updated: November 2025
    This function automatically opens the generated Excel workbook upon completion.
#>

}


 








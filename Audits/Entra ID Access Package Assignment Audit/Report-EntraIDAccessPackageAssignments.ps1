

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
        if ($InputObject) {
            $o = New-Object psobject
            foreach ($key in $InputObject.Keys) {
                $value = $InputObject[$key]
                if ($value -and $value.GetType().FullName -match 'System.Object\[\]') {
                    if ($value.Count -gt 0 -and $value[0] -is [hashtable]) {
                        Add-Member -InputObject $o -NotePropertyName $key -NotePropertyValue ($value | ConvertTo-PSCustomObject)
                    }
                    else {
                        Add-Member -InputObject $o -NotePropertyName $key -NotePropertyValue $value
                    }
                }
                elseif ($value -is [hashtable]) {
                    Add-Member -InputObject $o -NotePropertyName $key -NotePropertyValue (ConvertTo-PSCustomObject -InputObject $value)
                }
                else {
                    Add-Member -InputObject $o -NotePropertyName $key -NotePropertyValue $value
                }
            }
            Write-Output $o
        }
    }
    <#
.SYNOPSIS
    Converts a hashtable to a PowerShell custom object recursively.
#>
}



function igall {
    [CmdletBinding()]
    param (
        [string]$Uri,
        [switch]$Eventual,
        [int]$Limit = 1000
    )

    $nextUri = $uri
    $headers = @{ Accept = 'application/json' }
    if ($Eventual) { $headers['ConsistencyLevel'] = 'eventual' }

    $count = 0
    do {
        $result = Invoke-MgGraphRequest -Method GET -Uri $nextUri -Headers $headers
        $nextUri = $result.'@odata.nextLink'
        if ($result.value) {
            $result.value | ConvertTo-PSCustomObject
        }
        $count++
    } while ($nextUri -and ($count -lt $Limit))
    <#
.SYNOPSIS
    Retrieves all paginated results from a Microsoft Graph endpoint.

.PARAMETER Uri
    The Microsoft Graph API endpoint URI.

.PARAMETER Eventual
    Use the 'eventual' consistency level (for large data sets).

.PARAMETER Limit
    Maximum number of pages to retrieve.
#>
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
    Retrieves all user targets assigned to a specific Access Package.

.DESCRIPTION
    Queries Microsoft Graph for Entitlement Management assignments
    where the Access Package ID matches and state = 'Delivered'.

.PARAMETER AccessPackageId
    The ObjectId of the Access Package.

.REQUIRED_SCOPES
    EntitlementManagement.Read.All
    User.Read.All
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
    Compares a list of users against an Access Package's assigned users.
#>
}



function Get-UsersDynamic {
    [CmdletBinding(DefaultParameterSetName = "Filter")]
    param (
        [Parameter(ParameterSetName = "Filter")] [string]$EmployeeIdStartsWith,
        [Parameter(ParameterSetName = "Filter")] [string]$Country,
        [Parameter(ParameterSetName = "Filter")] [switch]$EmployeeIdNotNull,
        [Parameter(ParameterSetName = "Filter")] [switch]$EmployeeLeaveDateTimeNotNull,
        [Parameter(ParameterSetName = "Group")]  [string]$GroupId,
        [Parameter(ParameterSetName = "Manual")] [array]$Users
    )

    switch ($PSCmdlet.ParameterSetName) {
        "Filter" {
            $filterParts = @()
            if ($EmployeeIdStartsWith) { $filterParts += "startsWith(employeeId, '$EmployeeIdStartsWith')" }
            if ($Country) { $filterParts += "country eq '$Country'" }
            $filterParts += "accountEnabled eq true"

            $filterQuery = [string]::Join(' and ', $filterParts)
            $uri = "https://graph.microsoft.com/v1.0/users?`$filter=$filterQuery&`$select=id,displayName,userPrincipalName,employeeId,employeeHireDate,employeeLeaveDateTime,jobTitle,accountEnabled"
            
            Write-Host "🔎 Querying users with: $filterQuery" -ForegroundColor Yellow
            $data = igall $uri

            if ($EmployeeIdNotNull) { $data = $data | Where-Object { $_.employeeId } }
            if ($EmployeeLeaveDateTimeNotNull) { $data = $data | Where-Object { $_.employeeLeaveDateTime } }
        }

        "Group" {
            Write-Host "📂 Getting users from group ID: $GroupId" -ForegroundColor Yellow
            $uri = "https://graph.microsoft.com/v1.0/groups/$GroupId/members/microsoft.graph.user?`$select=id,displayName,userPrincipalName,employeeId,employeeHireDate,employeeLeaveDateTime,jobTitle,accountEnabled"
            $data = igall $uri
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
        }
    }
    <#
.SYNOPSIS
    Retrieves users dynamically using filters, group membership, or manual input.

.DESCRIPTION
    Allows flexible user retrieval for audits:
    - Filter by country, EmployeeId prefix, or enabled accounts.
    - Retrieve members from a group.
    - Manually specify user list.

.REQUIRED_SCOPES
    User.Read.All
    Group.Read.All
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
            $employeePrefix = Read-Host "EmployeeId starts with (or leave blank)"
            $useEmpIdNotNull = Read-Host "Filter EmployeeId not null? (y/n)"
            $useLeaveDateNotNull = Read-Host "Filter EmployeeLeaveDateTime not null? (y/n)"

            $params = @{}
            if ($country) { $params.Country = $country }
            if ($employeePrefix) { $params.EmployeeIdStartsWith = $employeePrefix }
            if ($useEmpIdNotNull -eq 'y') { $params.EmployeeIdNotNull = $true }
            if ($useLeaveDateNotNull -eq 'y') { $params.EmployeeLeaveDateTimeNotNull = $true }

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
}

function Export-AccessPackageReportToExcel {
    param(
        [Parameter(Mandatory)][PSObject[]]$CompareResults,
        [Parameter(Mandatory)][PSObject[]]$Summary,
        [Parameter(Mandatory)][PSObject[]]$TotalSummary,
        [Parameter(Mandatory)][string]$ExportPath
    )

    # --- Export DetailedReport ---
    $pkg = $CompareResults | Export-Excel -Path $ExportPath `
        -WorksheetName 'DetailedReport' `
        -TableStyle Medium2 -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow `
        -PassThru

    # --- Conditional formatting on Assigned column dynamically ---
    if ($CompareResults.Count -gt 0) {
        $detailedWs = $pkg.Workbook.Worksheets['DetailedReport']

        # Find the column number for "Assigned"
        $headerRow = 1
        $assignedCol = 0
        for ($col = 1; $col -le $detailedWs.Dimension.Columns; $col++) {
            if ($detailedWs.Cells[$headerRow, $col].Text -eq "Assigned") {
                $assignedCol = $col
                break
            }
        }

        if ($assignedCol -gt 0) {
            $startRow = 2
            $endRow = $CompareResults.Count + 1

            # Convert column number to Excel column letter dynamically (handles > Z)
            function Get-ExcelColumnLetter($colNumber) {
                $letter = ""
                while ($colNumber -gt 0) {
                    $colNumber--
                    $letter = [char](65 + ($colNumber % 26)) + $letter
                    $colNumber = [math]::Floor($colNumber / 26)
                }
                return $letter
            }

            $colLetter = Get-ExcelColumnLetter $assignedCol
            $address = "${colLetter}${startRow}:${colLetter}${endRow}"

            Add-ConditionalFormatting -Worksheet $detailedWs `
                -Address $address `
                -RuleType Equal -ConditionValue 'TRUE' -BackgroundColor LightGreen

            Add-ConditionalFormatting -Worksheet $detailedWs `
                -Address $address `
                -RuleType Equal -ConditionValue 'FALSE' -BackgroundColor LightPink
        }
    }

    # --- Export Summary sheet ---
    $Summary | Export-Excel -ExcelPackage $pkg `
        -WorksheetName 'Summary' `
        -AutoSize -AutoFilter -TableStyle Medium1 -TableName 'AccessPackageSummary'

    $summaryWs = $pkg.Workbook.Worksheets | Where-Object { $_.Name -eq 'Summary' }

    # --- Write totals below summary ---
    if ($summaryWs -and $TotalSummary.Count -gt 0) {
        $row = $Summary.Count + 4
        Set-ExcelRange -Worksheet $summaryWs -Range "A$row" -Value "Assignment Totals" -Bold -FontSize 12

        # --- Clean totals ---
        $cleanTotals = $TotalSummary | Select-Object AccessPackagePolicyName, Count | ForEach-Object {
            $props = $_.PSObject.Properties | Where-Object { $_.Value -ne $null -and $_.Value -ne "" }
            if ($props.Count -gt 0) {
                $newObj = New-Object PSObject
                foreach ($p in $props) {
                    $newObj | Add-Member -MemberType NoteProperty -Name $p.Name -Value $p.Value
                }
                $newObj
            }
        } | Where-Object { $_ -ne $null }

        if ($cleanTotals.Count -gt 0) {
            # --- Export totals table ---
            $tableName = "TotalsTable"
            $cleanTotals | Export-Excel -ExcelPackage $pkg -Worksheet 'Summary' `
                -StartRow ($row + 1) -StartColumn 1 -TableName $tableName -TableStyle Medium6 -Append

            # --- Calculate data range for chart ---
            $chartStartRow = $row + 2
            $chartEndRow = $chartStartRow + $cleanTotals.Count - 1

            # --- Add chart using explicit cell ranges ---
            Add-ExcelChart -Worksheet $summaryWs `
                -ChartType PieExploded3D `
                -XRange "A$chartStartRow:A$chartEndRow" `
                -YRange "B$chartStartRow:B$chartEndRow" `
                -Title "Assignment Overview" `
                -Top (($row - 1) * 15) `
                -Left 300 -Width 400 -Height 300 | Out-Null
        }
    }

    # --- Reopen Excel package if reference lost ---
    if (-not $pkg -or -not $pkg.Workbook) {
        if (Test-Path $ExportPath) {
            try {
                Write-Host "Reopening Excel package from disk before save..." -ForegroundColor Yellow
                $pkg = Open-ExcelPackage -Path $ExportPath
            }
            catch {
                Write-Host "⚠️ Failed to reopen Excel package at $ExportPath" -ForegroundColor Red
            }
        }
    }

    # --- Final Save ---
    try {
        if ($pkg -and $pkg.Workbook) {
            $pkg.Save()
            Close-ExcelPackage $pkg -Show
            Write-Host "✅ Excel report created successfully: $ExportPath" -ForegroundColor Green
            Write-Host "Sheets included: Summary, DetailedReport" -ForegroundColor Cyan
        }
        else {
            throw "ExcelPackage not valid."
        }
    }
    catch {
        Write-Host "❌ Export failed — ExcelPackage not valid or save error." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Yellow
    }
}

 








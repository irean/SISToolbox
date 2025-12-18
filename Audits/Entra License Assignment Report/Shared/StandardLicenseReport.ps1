. (Join-Path $PSScriptRoot \licenses.ps1)
. (join-Path $PSScriptRoot \companies.ps1)

Class CompanyNames : System.Management.Automation.IValidateSetValuesGenerator {
    [string[]] GetValidValues() {

        return [string[]] $Global:tenants.Keys
    }
}
function test-module {
    [CmdletBinding()]
    param(
        [String]$Name

    )
    Write-Host "────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "Checking module $name. Because apparently we can't assume anything works." -ForegroundColor Cyan
    if (-not (Get-Module $Name)) {
       Write-Host "📦 Module $Name not imported — classic. Don’t worry Jens, I’ll fix it." -ForegroundColor Yellow
        try {
            if ($Name -eq 'Microsoft.Graph') {
                Write-Host "⏳ Importing Microsoft.Graph… Jens, this is where your patience gets tested." -ForegroundColor Magenta
                Import-Module $Name  -ErrorAction Stop
            }
            elseif ($Name -eq 'Az') {
                Write-Host "⏳ Importing Az… yes, *that* Az. Go get coffee, Jens, you’ve earned it." -ForegroundColor Magenta
            }
            else {
                  Write-Host "📥 Importing module $Name. Jens, if this works on the first try, treat yourself to a pastry." -ForegroundColor DarkCyan
                Import-Module $Name  -ErrorAction Stop
            }

        }
        catch {
             Write-Host "❌ Module $Name not found. Installing it now — because planning is for other people." -ForegroundColor Red
            Install-Module $Name -Scope CurrentUser -AllowClobber -Force -AcceptLicense -SkipPublisherCheck
             Write-Host "📥 Importing freshly installed module $Name. Jens, pretend this was step one." -ForegroundColor DarkCyan

            Import-Module $Name -ErrorAction stop

        }
    }
    else {
        Write-Host "✅ Module $Name is already imported. Miracles do happen, Jens." -ForegroundColor Green
    }
     Write-Host "────────────────────────────────────────" -ForegroundColor DarkGray
}
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
Test-Module -name Microsoft.Graph.Users
test-Module -name Microsoft.Graph.Authentication
Test-module -name Microsoft.Graph.DeviceManagement
Test-Module -name ImportExcel
function Connect-LicenseReport {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $tenantId
    )
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
    }
    catch {
        <#Do this if a terminating exception happens#>
    }
Write-Host "Connecting to tenant $tenantId. Jens, if this fails, just look disappointed and nod knowingly." -ForegroundColor Cyan

    Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All" , "DeviceManagementConfiguration.Read.All", "DeviceManagementManagedDevices.Read.All" , "Organization.Read.All", "AuditLog.Read.All" -TenantId $tenantId  -ContextScope Process -ErrorAction Stop
    if (-not (Get-Mgcontext)) {
        Throw "Failed to connect to $tenantID. Add this to the list of things we blame on Microsoft."
    }
}
function Connect-MGGraphAPI {
    param (
        [Parameter()]
        [String]$clientID,
        [Parameter()]
        [String]$tenantID
    )

Write-Host "Authenticating with ClientID $clientID. Jens, try not to type the Client Secret into Teams." -ForegroundColor Yellow

    $secret = Get-Credential -Credential $clientId
    Connect-MgGraph -TenantId $tenantId -ClientSecretCredential $secret

}



function Get-LicenseReport {
    param(
        [Parameter(Mandatory = $false)]
        [ValidateSet(
            [CompanyNames]
        )]
        [String]$organization
    )
    Write-Host "Starting license report. Because this script still refuses to run itself." -ForegroundColor Cyan
    if ($organization) {
        Write-Host "Generating report for: $organization. Don’t worry, Jens, this part almost never fails." -ForegroundColor Cyan
        Get-SingleLicenseReport -Name $organization -clientId $tenants[$organization]['clientID'] -tenantid $tenants[$organization]['TenantID'] -Extensions $tenants[$organization]['extensions']
    }
    else {
        Write-Host "No organization provided. Cool. Guess we're doing *everyone* today." -ForegroundColor Yellow
        foreach ($tenant in $tenants.Keys) {
            Write-Host "----> Processing tenant: $tenant. If this takes long, blame the cloud." -ForegroundColor Cyan
            Get-SingleLicenseReport -Name $tenant -clientId $tenants[$tenant]['clientID'] -tenantid $tenants[$tenant]['tenantID'] -Extensions $tenants[$tenant]['extensions']

        }
    }
    Write-Host "License report completed. Go reward yourself by pretending this was hard." -ForegroundColor Green
}

function Get-SingleLicenseReport {
    param (
        [Parameter()]
        [String]$name,
        [Parameter()]
        [String]$tenantid,
        [parameter()]
        [String]$clientId,
        [String[]]$Extensions

    )

    Write-Host "Connecting to tenant '$name'. If this hangs, it's definitely Microsoft's fault." -ForegroundColor Green
    try {

        Connect-MgGraphAPI -tenantId $tenantID -clientID $clientId
        Write-Verbose "Connected to Graph API. Somehow."

        $skuMapping = @{}
        Write-Verbose "Initialized SKU mapping dictionary, because Microsoft thinks product names are optional."

        function convertto-license {
            [CmdletBinding()]
            param (
                $User,
                [string]$License,
                [string]$sku,
                [string]$Country,
                [string[]]$extensions
            )
            Write-Verbose "Converting user $($User.UserPrincipalName) to license object. Try not to look excited."
            $res = [Ordered]@{
                User                             = $User.UserPrincipalName
                EmployeeID                       = $User.employeeID
                SKU                              = $sku
                License                          = $License
                Name                             = $User.DisplayName
                Title                            = $User.JobTitle
                City                             = $User.City
                StreetAddress                    = $User.StreetAddress
                Country                          = $user.Country
                Company                          = $User.CompanyName
                Department                       = $User.Department
                USertype                         = $User.UserType
                Enabled                          = $user.AccountEnabled
                Created                          = $User.createdDateTime
                LastSignInDateTime               = $User.signInActivity.lastSignInDateTime
                lastNonInteractiveSignInDateTime = $User.signInActivity.lastNonInteractiveSignInDateTime

            }
            if ($extensions) {
                Write-Verbose "Adding extension attributes for $($User.UserPrincipalName). Because apparently default attributes weren’t enough."
                foreach ($prop in $extensions) {
                    $res.add($prop, $user.$prop)
                }
            }
            return [PSCustomObject]$res

        }
        $Report = [System.Collections.Generic.List[Object]]::new()
        Write-Information "Created base report object for $name. It’s empty, like our hopes."
        $fields = @(
            'Displayname',
            'UserPrincipalName',
            'EmployeeID',
            'EmployeeNumber',
            'JobTitle',
            'City',
            'Department',
            'CompanyName',
            'StreetAddress',
            'Country',
            'UserType',
            'onPremisesExtensionAttributes',
            'AdditionalProperties',
            'createdDateTime',
            'AccountEnabled',
            'assignedLicenses',
            'assignedLicenses',
            'signInActivity',
            'lastNonInteractiveSignInDateTime',
            'lastSignInDateTime'
        )
        if ($extensions) {
            Write-Verbose "Adding custom extension fields: $($extensions -join ', '). Because why be simple."
            $fields += $Extensions
        }
        $f = $fields -join '%2c'

        $uri = "https://graph.microsoft.com/beta/users?`$top=100&`$Select=$f"
        Write-Host "Retrieving users for '$name'... Brace yourself, Jens." -ForegroundColor Cyan
      
        do {
            Write-Verbose "Fetching user batch from Graph: $uri"
            $response = Invoke-MgGraphRequest -Method GET -Uri $uri
            $userlist = $response.value
            $uri = $response.'@odata.nextlink'
            Write-Host "Processing $($userList.Count) users... because someone has to." -ForegroundColor Yellow
            $index = 0
            ForEach ($User in $Userlist) {
                Write-Progress -Activity "Processing Users in $name" -PercentComplete ($index / $userlist.count * 100) -Status $user.userprincipalname
                $index++
                Write-Verbose "Processing $($User.UserPrincipalName). Hopefully they still work here."

                ForEach ($skuId in $User.assignedLicenses.skuId) {
                    if (-not $skuMapping[$skuId]) {
                        Write-Verbose "Fetching SKU details for $skuId. Because Microsoft loves GUIDs more than humans."
                        $details = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/users/$($user.id)/licenseDetails" | Select-Object -ExpandProperty value
                        foreach ($detail in $details) {
                            Write-Verbose "Mapped SKU $($detail.SkuId) → $($detail.SkuPartNumber). You're welcome, Jens."
                            $skuMapping[$detail.SkuId] = $detail.SkuPartNumber
                        }
                    }
                    $license = Get-LicenseFriendlyName -sku $skuMapping[$SkuId]
                    Write-Verbose "Translated SKU to license: $license."

                    $ReportLine = convertto-license -User $User -License $License -Country $Country   -Sku $skuMapping[$SkuId] -extensions $Extensions
                    $Report.Add($ReportLine)
                }

                if ((-not $User.assignedLicenses) -or $User.assignedLicenses.count -eq 0) {
                    Write-Verbose "$($User.UserPrincipalName) has no licenses. Must be having a great time."

                    $ReportLine = convertto-license -User $User -License "" -Country $Country -sku "" -extensions $Extensions
                    $Report.Add($ReportLine)
                }
            }
        } while ($uri)
        Write-Host "Finished processing users for $name. All fingers still attached." -ForegroundColor Green
        Write-Information "Creating license summary. It's exactly as thrilling as it sounds."


        $Groupdata = $Report | Group-Object -Property License | Sort-Object Count -Descending | Select-Object Name, Count
        $GroupData
        Write-Host "Generating Excel report for '$name'. Please don’t close Excel this time, Jens." -ForegroundColor Cyan

        # Set sort properties so that we get ascending sorts for one property after another
        $date = Get-Date -Format yyyyMMdd
        $year = Get-Date (Get-Date).AddMonths(1)  -format yyyy 
        $month = Get-Date (Get-Date).AddMonths(1) -format yyyy-MM
        $orgname = $name
        $outDir = "$env:USERPROFILE\Documents\Reports\$orgname\$year\$month"
        if (Test-Path "$env:USERPROFILE\<synced sharepoint folder>") {
            $outDir = "$env:USERPROFILE\<synced sharepoint folder>\$orgname\$year\$month"
        }
        else {
            Write-Information "$outdir\$($orgname)LicenseUser-$Date.xlsx"
        }


        $pivotLicense = $Report | Sort-Object -Property User, LIcense, Country  | Where-Object {
            $_.License
        }
        $report2 = $Report | Select-object Name, User, Sku | Where-object {
            ($_.Sku -match 'SPE' -and $_.Sku -notmatch '_SEC') -or $_.Sku -match 'exchange' -or ($_.sku -match 'Business' -and $_.Sku -notmatch 'DYN365') -or $_.sku -match 'pack'
        }
        if (test-path "$outdir\$($orgname)LicenseUser-$Date.xlsx" ) {
            Remove-Item "$outdir\$($orgname)LicenseUser-$Date.xlsx"
        }

        $date30 = (Get-Date).AddDays(-30)

        if ($orgname -match 'Marker') {

            $devices = IGall "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices"  | Where-Object {


                <#$_.lastSyncDateTime -ge $date30 -and#>$_.deviceName -notmatch 'SKOE' -and $_.deviceName -notmatch 'EIPAD' -and $_.userPrincipalName -notmatch 'sandra.saluti' -and $_.userPrincipalName -notmatch 'gean' -and $_.userPrincipalName -notmatch 'geir.anger' -and $_.userPrincipalName -notmatch 'geir.anger' -and $_.model -notmatch "Virtual Machine"
            } | Select-Object id, deviceName, managedDeviceOwnerType, operatingSystem, lastSyncDatetime, compliancestate , osversion
        }
        else {
            $devices = IGall "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices"  | Where-Object {


                <#$_.lastSyncDateTime -ge $date30 -and#>$_.userPrincipalName -notmatch 'sandra.saluti' -and $_.userPrincipalName -notmatch 'gean' -and $_.userPrincipalName -notmatch 'geir.anger' -and $_.userPrincipalName -notmatch 'geir.anger' -and $_.model -notmatch "Virtual Machine"
            } | Select-Object id, deviceName, managedDeviceOwnerType, operatingSystem, lastSyncDatetime, compliancestate , osversion

        }

        $pivotM365Licenses = $PivotLicense #| Where-Object {
        #    $_.License -match 'Microsoft 365 E3'
        # }
        $pivotDevices = $devices | Sort-Object -Property operatingSystem, deviceEnrollmentType, id, managedDeviceOwnerType

        $excel = $Report | Export-Excel -Path "$outdir\$($orgname)LicenseUser-$Date.xlsx" -KillExcel -WorksheetName "UserData" -AutoSize -BoldTopRow -TableName "users" -TableStyle Medium6 -FreezeTopRow -PassThru -ClearSheet `
            #  -IncludePivotTable -PivotRows "Country" -PivotColumns "LicenseGroup" -PivotData @{"User" = "Count" }

        $excel = $PivotLicense | Export-Excel -ExcelPackage $excel -KillExcel -WorksheetName "License" -AutoSize -BoldTopRow -TableName "Users by License"  `
            -TableStyle Medium6 -PassThru -ClearSheet `
            -IncludePivotTable -PivotRows "License" -PivotColumns "Country" -PivotData @{"User" = "Count" }
        $excel = $pivotM365Licenses | Export-Excel -ExcelPackage $excel -KillExcel -WorksheetName "LicenseByCompany" -AutoSize -BoldTopRow -TableName "License By Company"  `
            -TableStyle Medium6 -PassThru -ClearSheet `
            -IncludePivotTable -PivotRows "Company" -PivotColumns "License" -PivotData @{"User" = "Count" }
        $excel = $availLicenses | Export-Excel -ExcelPackage $excel -KillExcel -WorksheetName "LicenseActivated"  -AutoSize -BoldTopRow -TableName "Unused Licenses" `
            -TableStyle Medium6 -StartColumn 11 -PassThru -ClearSheet

        $excel = $report2  | Export-Excel -ExcelPackage $excel -KillExcel -WorksheetName "MailboxUsers" -TableName "Users with enabled mailboxes"  -AutoSize `
            -TableStyle Medium6 -PassThru -ClearSheet
        $excel = $Pivotdevices  | Export-Excel -ExcelPackage $excel -KillExcel -WorksheetName "Devices" -TableName "Devices"  -AutoSize `
            -TableStyle Medium6 -PassThru -ClearSheet `
            -IncludePivotTable -PivotRows "operatingSystem" -PivotColumns "managedDeviceOwnerType" -PivotData @{"id" = "Count" } -PivotTableName 'Device Count'




        Close-ExcelPackage $excel -Show

        Write-Host "Excel report for '$name' completed. Frame it, laminate it, cherish it." -ForegroundColor Green

    }
    catch {
        Write-Error "Something broke while processing $name. Probably not your fault, Jens, but let's pretend it is: $_"
    }

}

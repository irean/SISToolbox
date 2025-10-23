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
        elseif($result.value.count -eq 0){
            @()
        }
        elseif ($result) {
            $result | ConvertTo-PSCustomObject
        }
        $count += 1
    } while ($nextUri -and ($count -lt $limit))
}

function ig {
    [CmdletBinding()]
    param (
        [string]$Uri
    )
    $result = Invoke-MgGraphRequest -Method GET -uri $uri
    if ($result.value) {
        $result.value | ConvertTo-PSCustomObject
    }
    else {
        $result | ConvertTo-PSCustomObject
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
          if ($value.Count -gt 0 -and $value[0].GetType().FullName -match  'System.Collections.Hashtable') {
            $tempVal = $value | ConvertTo-PSCustomObject
            Add-Member -InputObject $o -NotePropertyName $key -NotePropertyValue $tempVal
          } elseif ($value.Count -gt 0 -and $value[0].GetType().FullName -match  'System.String') {
            $tempVal = $value | ForEach-Object { $_ }
            Add-Member -InputObject $o -NotePropertyName $key -NotePropertyValue $tempVal
          }
        } elseif ($value -and $value.GetType().FullName -match 'System.Collections.Hashtable') {
          Add-Member -InputObject $o -NotePropertyName $key -NotePropertyValue (ConvertTo-PSCustomObject -InputObject $value)
        } else {
          Add-Member -InputObject $o -NotePropertyName $key -NotePropertyValue $value
        }
      }

      Write-Output $o
    }
  }
}

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
# Disconnecting any existing Microsoft Graph session
Write-Host "🔌 Disconnecting previous Microsoft Graph sign-in..." -ForegroundColor Yellow
Disconnect-MgGraph

Write-Host "🔎 Checking module: Microsoft.Graph.Authentication..." -ForegroundColor Blue
Test-Module -Name Microsoft.Graph.Authentication
Write-Host "🔎 Checking module: ImportExcel..." -ForegroundColor Blue
Test-Module -Name ImportExcel 

Write-Host "✅ All required modules are available. Please connect to Microsoft Graph." -ForegroundColor Green


Connect-Mggraph -Scopes "Organization.Read.All", "user.read.all", "groupmember.read.all", "group.read.all", "application.read.all"

$date = Get-Date -format yyyy-MM-dd 
#Get organization name and remove any spaces in name
$org = (Igall "https://graph.microsoft.com/v1.0/organization" | Select-Object -ExpandProperty displayName) -replace ' ', ''

#Select  Data FilePath
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# Ask the user if they have an Excel list of users
Write-Host ""
Write-Host "💡 When using an Excel file to target specific users," -ForegroundColor Cyan
Write-Host "   ensure it contains a column with the header 'userPrincipalName' — this is required to identify each user." -ForegroundColor Cyan
Write-Host ""


$hasExcelList = Read-Host "Do you have an Excel list of users? (Y/N)" 

if ($hasExcelList -match '^(Y|y)$') {

    Write-Host "Opening file picker for your user data file... If you don't see it, check behind other windows!" -ForegroundColor Yellow
    
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
    $FileBrowser.Filter = "All files (*.*)|*.*"
    $null = $FileBrowser.ShowDialog()

    $FilePath = $FileBrowser.FileName
    Write-Host "Selected file path: $FilePath" -ForegroundColor Green
    #Get users from exel 

    $users = Import-Excel -path $filepath
    Write-Host "Retrieving all Entra ID for listed userPrincipalNames..." -ForegroundColor Cyan

    $eusers = $users | ForEach-Object {
        $displayname = $_.Displayname
        $upn = $_.userPrincipalName

        $user = igall "https://graph.microsoft.com/v1.0/users/$upn"

        if ($user.userPrincipalName) {
            $user
        }
        else {
            write-host "Could not find $displayname" -ForegroundColor Red
        }
    }
}
elseif ($hasExcelList -match '^(N|n)$') {
    Write-Host "Okay, no Excel list will be used and all users will be targeted"
    $progress = 0
    $eusers = igall "https://graph.microsoft.com/v1.0/users?`$filter=userType eq 'member'" | ForEach-Object {
        $progress++
        if ($progress % 100 -eq 0) {
            Write-Progress -Activity "Fetching users" -Status "Retrieved $progress users..." -PercentComplete (($progress % 10000) / 100)
        }
        $_
    }
    Write-Progress -Activity "Fetching users" -Completed
    Write-Host "✅ Completed. Retrieved $($eusers.Count) users." -ForegroundColor Green
}
else {
    Write-Host "Invalid input. Please run the script again and answer Y or N." -ForegroundColor Red
}


#save file location
Write-Host "Opening picker for save folder path.. if you don't see it, again, check behind other windows!" -ForegroundColor Yellow
Add-Type -AssemblyName System.Windows.Forms
$FileBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop') 
}
if ($FileBrowser.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host "No folder selected"
    return
}
$folder = $FileBrowser.SelectedPath

$savepath = "$($folder)\$($org)_Application_Dynamicgroup_Assignments_$($date).xlsx"




Write-Host "Selected SavePath $savepath" -ForegroundColor Green
Write-Host "Searching through all appRoleAssignments for users..." -ForegroundColor Cyan
$appcache = @{}

$assignments = $eusers | ForEach-Object {

    $auser = $_

    $upn = $_.userPrincipalName
    $a = Igall "https://graph.microsoft.com/v1.0/users/$upn/appRoleAssignments" 
    if($a.count -gt 0){
    $a | ForEach-Object {
        $assignmentObj = $auser.PSObject.Copy() |
        Add-Member -NotePropertyName assignmentType -NotePropertyValue $_.principalType -Force -PassThru |
        Add-Member -NotePropertyName GroupDisplayName -NotePropertyValue $_.principalDisplayName -Force -PassThru |
        Add-Member -NotePropertyName resourceDisplayName -NotePropertyValue $_.resourceDisplayName -Force -PassThru |
        Add-Member -NotePropertyName resourceID -NotePropertyValue $_.resourceId -Force -PassThru

        # get the service principal for this resource
        if (-not $appcache[$_.resourceId]) {
            Write-Host "Service principal not cached, fetching it $($_.resourceDisplayName) $($_.resourceId)" -ForegroundColor Cyan
            $spTmp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($_.resourceId)"
            try {
                $spScim = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($_.resourceId)/synchronization/jobs"
                Add-Member -InputObject $spTmp -NotePropertyName scim -NotePropertyValue $spScim.value -Force
            }
            catch {}
            $appcache.Add($_.resourceId, $spTmp)
        }
        $sp = $appcache[$_.resourceId]
        if ($sp.preferredSingleSignOnMode) {
            $assignmentObj | Add-Member -NotePropertyName SSO -NotePropertyValue $sp.preferredSingleSignOnMode -Force

        }
        else {
            if (-not $appcache[$sp.appId]) {
                Write-Host "app not cached, fetching it $($sp.appId)" -ForegroundColor Cyan
                $appcache.Add($sp.appId, (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/applications?`$filter=appId eq '$($sp.appId)'"))
            }
            $app = $appcache[$sp.appId]
            if ($app.value.count -eq 0) {
                $assignmentObj | Add-Member -NotePropertyName SSO -NotePropertyValue False -Force
            }
            else {
                $appObj = $app.value[0]
                Write-Host "Skipping '$($appObj.DisplayName)' — this application is already in the cache." -ForegroundColor Yellow

                $uris = @()
                if ($appObj.web.redirectUris) { 
                    $uris += $appObj.web.redirectUris 
                }
                if ($appObj.spa.redirectUris) { 
                    $uris += $appObj.spa.redirectUris 
                }
                if ($appObj.publicClient.redirectUris) {
                    $uris += $appObj.publicClient.redirectUris 
                }

                if ($uris -match "oidc" -or $uris -match "signin" -or $appObj.requiredResourceAccess.Count -gt 0) {
                    $assignmentObj | Add-Member -NotePropertyName SSO -NotePropertyValue 'oidc' -Force
        
                }
                else {
                    $assignmentObj | Add-Member -NotePropertyName SSO -NotePropertyValue False -Force
                }

            }
        }
        try {
            if ($sp.scim.schedule.state -eq 'Active') {

                $assignmentObj | Add-Member -NotePropertyName SCIM -NotePropertyValue True -Force
            }
            else {
                $assignmentObj | Add-Member -NotePropertyName SCIM -NotePropertyValue False -Force 
            }
        }
        catch {
            $assignmentObj | Add-Member -NotePropertyName SCIM -NotePropertyValue False -Force

        }
        
        $assignmentObj
       
        
    }
}
}

$assignemntsorted = $assignments | Sort-Object userPrincipalName -Descending | Sort-Object userPrincipalName -Descending | Select-Object UserPrincipalName, DisplayName, officeLocation, jobTitle, assignmentType, GroupDisplayName, resourceDisplayName, resourceID, SSO, SCIM

Write-Host "✅ App role assignments checked — proceeding to next step." -ForegroundColor Green
Write-Host "→ Checking user memberships in dynamic groups..." -ForegroundColor Cyan

$groups = $eusers | Foreach-Object {
    $upn = $_.userPrincipalName
    $guser = $_

    Write-Host "Checking Entra ID group memberships for user '$upn'..." -ForegroundColor Cyan

    $g = Igall "https://graph.microsoft.com/beta/users/$upn/memberof" | Where-Object { 
        $_.grouptypes -contains 'DynamicMembership' 
    }

    $g | Foreach-Object {
Write-Host "Fetching dynamically assigned group '$($_.DisplayName)'..." -ForegroundColor Cyan

        add-member -inputObject $guser.psobject.copy() -NotePropertyName GroupDisplayName -NotePropertyValue $_.displayname -Force -PassThru |
        add-member -NotePropertyName groupTypes -NotePropertyValue $_.groupTypes -force -PassThru |
        add-member -NotePropertyName GroupDescription -NotePropertyValue $_.description -force -PassThru |
        add-member -NotePropertyName mailenabled -NotePropertyValue $_.mailenabled -force -PassThru |
        add-member -NotePropertyName Securityenabled -NotePropertyValue $_.Securityenabled -force -PassThru |
        add-member -NotePropertyName membershipRule -NotePropertyValue $_.membershipRule -force -PassThru 
    }
}
$groupssorted = $groups | Sort-object userprincipalName -Descending | Select-Object UserPrincipalName, DisplayName, officeLocation, jobTitle, GroupDisplayName, groupTypes, GroupDescription, mailenabled, Securityenabled, membershipRule 
Write-Host "All dynamically assigned groups have been captured. Saving output file..." -ForegroundColor Green

$assignemntsorted | Export-Excel -path $savepath -TableStyle Medium2 -WorksheetName ApplicationAssignments -AutoSize
$groupssorted | Export-Excel -path $savepath -WorksheetName DynamicGroupAssignments -TableStyle Medium3  -AutoSize -show
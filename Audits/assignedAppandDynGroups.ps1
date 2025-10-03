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

Write-Host "Opening file picker for your user data file... If you don't see it, check behind other windows!" -ForegroundColor Yellow
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
$FileBrowser.Filter = "All files (*.*)|*.*"
$null = $FileBrowser.ShowDialog()

$FilePath = $FileBrowser.FileName
Write-Host "Selected file path: $FilePath" -ForegroundColor Green

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


#Get users from exel 

$users = Import-Excel -path $filepath

$eusers = $users | ForEach-Object {
    $displayname = $_.Displayname
    $upn = $_.userprincipalname

    $user = igall "https://graph.microsoft.com/v1.0/users?`$filter=imAddresses/any(i:i eq '$upn')"

    if ($user.userprincipalName) {
        $user
    }
    else {
        write-host "Could not find $displayname"
    }
}

$assignments = $eusers | Foreach-Object {

    $auser = $_

    $upn = $_.userPrincipalName
    $a = Igall "https://graph.microsoft.com/v1.0/users/$upn/appRoleAssignments" 
    $a | Foreach-Object {
       
        add-member -InputObject $auser.psobject.copy() -NotePropertyName assignmentType -NotePropertyValue $_.principalType -Force -PassThru |
        add-member -NotePropertyName  GroupDisplayName -NotePropertyValue $_.principalDisplayName -force -PassThru |
        add-member -NotePropertyName resourceDisplayName -NotePropertyValue $_.resourceDisplayName -force -PassThru |
        add-member -NotePropertyName resourceID -NotePropertyValue $_.resourceId -force -PassThru
    }

    
}

$assignemntsorted = $assignments | Sort-Object userPrincipalName -Descending | Sort-Object userPrincipalName -Descending | Select-Object UserPrincipalName, DisplayName, officeLocation, jobTitle, assignmentType, GroupDisplayName, resourceDisplayName, resourceID, SSO, Provisioning 

#Getting all groupmemeberships for user

$groups = $eusers | Foreach-Object {
    $upn = $_.userPrincipalName
    $guser = $_

    $g = Igall "https://graph.microsoft.com/beta/users/$upn/memberof" | Where-Object { 
        $_.grouptypes -contains 'DynamicMembership' 
    }
    $g | Foreach-Object {
        add-member -inputObject $guser.psobject.copy() -NotePropertyName GroupDisplayName -NotePropertyValue $_.displayname -Force -PassThru |
        add-member -NotePropertyName groupTypes -NotePropertyValue $_.groupTypes -force -PassThru |
        add-member -NotePropertyName GroupDescription -NotePropertyValue $_.description -force -PassThru |
        add-member -NotePropertyName mailenabled -NotePropertyValue $_.mailenabled -force -PassThru |
        add-member -NotePropertyName Securityenabled -NotePropertyValue $_.Securityenabled -force -PassThru |
        add-member -NotePropertyName membershipRule -NotePropertyValue $_.membershipRule -force -PassThru 
    }
}
$groupssorted = $groups | Sort-object userprincipalName -Descending | Select-Object UserPrincipalName, DisplayName, officeLocation, jobTitle, GroupDisplayName, groupTypes, GroupDescription, mailenabled, Securityenabled, membershipRule 

$assignemntsorted | Export-Excel -path $savepath -TableStyle Medium2 -WorksheetName ApplicationAssignments -AutoSize
$groupssorted | Export-Excel -path $savepath -WorksheetName DynamicGroupAssignments  -show
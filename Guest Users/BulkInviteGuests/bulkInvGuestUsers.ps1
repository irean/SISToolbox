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

function Select-FolderPath {
    [CmdletBinding()]
    param()



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
function Select-FilePath {
    [CmdletBinding()]
    param(
        [string]$Title = "Select a file",
        [string]$Filter = "All files (*.*)|*.*"
    )

    Add-Type -AssemblyName System.Windows.Forms

    $FileDialog = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        Title            = $Title
        Filter           = $Filter
        InitialDirectory = [Environment]::GetFolderPath("Desktop")
        Multiselect      = $false
    }

    $form = New-Object System.Windows.Forms.Form -Property @{ TopMost = $true }
    $result = $FileDialog.ShowDialog($form)

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $file = $FileDialog.FileName
        Write-Host "✅ File selected: $file" -ForegroundColor Green
        return $file
    }
    else {
        Write-Host "❌ No file selected. Exiting script." -ForegroundColor Red
        return $null
    }

    <#
    .SYNOPSIS
        Opens a file picker dialog.

    .DESCRIPTION
        Displays a Windows file selection dialog and returns the selected file path.
        The dialog is forced to the top of the screen.

    .EXAMPLE
        $filePath = Select-FilePath

    .EXAMPLE
        $filePath = Select-FilePath -Filter "CSV files (*.csv)|*.csv"

    .RETURNS
        [string] The selected file path, or $null if cancelled.
    #>
}
Write-Host "Starting guest onboarding script..." -ForegroundColor Cyan

Write-Host "Checking Microsft Graph Authentication"
Test-Module -name Microsoft.Graph.Authentication
Write-Host "Checking module Import Excel"
test-module -name ImportExcel 

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan

connect-mggraph -scope "user.invite.all", "user.readwrite.all", "Organization.Read.All"
Write-Host "Checking Microsoft Graph connection context..." -ForegroundColor Cyan
$c = Get-MgContext
if ($c) {
    Write-Host "Microsoft Graph context found." -ForegroundColor Green
    Write-Host "Retrieving organization information..." -ForegroundColor Cyan
    $o = Igall "https://graph.microsoft.com/v1.0/organization"
    $dipslayname = $o.displayName

    Write-Host "Connected Organization: $dipslayName" -ForegroundColor Yellow
    Write-Host "Prompting for tenant confirmation..." -ForegroundColor Cyan

    $choice = Read-Host "You are connected to $dipslayName, do you want to add your guest here Y/Yes| N/No"
    if ($choice -match '^(y|yes)$') {
        Write-Host "User confirmed correct organization. Continuing..." -ForegroundColor Green
        Write-Host "Opening file picker for Excel input..." -ForegroundColor Cyan

        $filepath = Select-FilePath -Title "Select excelfil with user emails" -Filter "Excel (*.xlsx)|*.xlsx"
        Write-Host "Excel file selected: $filepath" -ForegroundColor Green

    }
    else {
        Write-Host "User chose not to continue. Disconnecting from Microsoft Graph." -ForegroundColor Yellow
        Disconnect-MgGraph
        return
        

    }
}
else {
    Write-Host "No Microsoft Graph context found. Script cannot continue." -ForegroundColor Red
}


#Import all users
Write-Host "Importing users from Excel file..." -ForegroundColor Cyan
$excelUsers = Import-Excel -path $filepath
Write-Host "Excel import completed." -ForegroundColor Green


$excelUsers | ForEach-Object {

    $excelUser = $_
    $mail = $_.mail
   
    Write-Host "----------------------------------------" -ForegroundColor DarkGray
    Write-Host "Processing user: $mail" -ForegroundColor Cyan

    # Build body
    Write-Host "Building request body from Excel attributes..." -ForegroundColor DarkGray
    $exclude = @('mail')
    $body = @{}

    foreach ($prop in $excelUser.PSObject.Properties) {
        if ($exclude -notcontains $prop.Name -and
            $null -ne $prop.Value -and
            $prop.Value -ne '') {
            Write-Host "   Adding attribute '$($prop.Name)'" -ForegroundColor DarkGray
            $body[$prop.Name] = $prop.Value
        }
    }
    Write-Host "Body contains $($body.Count) attributes" -ForegroundColor DarkGray
 Add-Member -InputObject $excelUser -NotePropertyName Status -NotePropertyValue 'Processing' -Force
    Add-Member -InputObject $excelUser -NotePropertyName ErrorStep -NotePropertyValue $null -Force
    Add-Member -InputObject $excelUser -NotePropertyName ErrorMessage -NotePropertyValue $null -Force
    Add-Member -InputObject $excelUser -NotePropertyName UserId -NotePropertyValue $null -Force

    # Check if user exists
    Write-Host "Checking if user exists in tenant..." -ForegroundColor DarkGray
    $existingUser = igall "https://graph.microsoft.com/v1.0/users?`$filter=mail eq '$mail'" 
    

    if ($existingUser.mail -and $existingUser.Count -gt 0 -and $existingUser[0].id) {

        $id = $existingUser[0].id
        $excelUser.UserId = $id
        $excelUser.Status = 'Updated'
        Write-Host "User already exists. UserId: $id" -ForegroundColor Yellow
    }
    else {
        Write-Host "User not found. Sending invitation..." -ForegroundColor Yellow

        $params = @{
            invitedUserEmailAddress = $mail
            inviteRedirectUrl       = "https://portal.office.com"
        }
        Write-Host "Inviting guest user..." -ForegroundColor DarkGray

        try {
            $guest = Invoke-MgGraphRequest `
                -Method POST `
                -Uri "https://graph.microsoft.com/v1.0/invitations" `
                -ContentType "application/json" `
                -Body $params

            $id = $guest.invitedUser.id
            $excelUser.UserId = $id
            $excelUser.Status = 'Invited'
            $wasInvited = $true
            Write-Host "Invitation sent. Guest userId: $id" -ForegroundColor Yellow
        }
        catch {
            $excelUser.Status = 'Failed'
            $excelUser.ErrorStep = 'Invite'
            $excelUser.ErrorMessage = $_.Exception.Message
            return
        }


    }

    # Retry ONLY if user was invited
    if ($wasInvited) {
        Write-Host "Waiting for invited user to become available in directory..." -ForegroundColor DarkGray

        $max = 10
        $delay = 3

        for ($i = 1; $i -le $max; $i++) {
            try {
                Write-Host "Attempt $i of $max to resolve user..." -ForegroundColor DarkGray
                $graphUser = igall "https://graph.microsoft.com/v1.0/users/$id"
                if ($graphUser) {
                    Write-Host "User is now available after $i attempts" -ForegroundColor Green
                    break
                }
            }
            catch {
                Write-Host "User not yet available. Waiting $delay seconds..." -ForegroundColor DarkGray
                Start-Sleep -Seconds $delay
            }
        }
    }
    if (-not $id) {
        Write-Host "❌ No user ID resolved for $mail – skipping update" -ForegroundColor Red
        $excelUser.Status = 'Skipped'
        $excelUser.ErrorStep = 'Retry'
        $excelUser.ErrorMessage = 'User not available after invitation'
        return

        
    }

    # Update (always happens) 
    Write-Host "Updating user with ID: $id" -ForegroundColor Cyan

    try {
            
        Invoke-MgGraphRequest `
            -Method PATCH `
            -Uri "https://graph.microsoft.com/v1.0/users/$id" `
            -ContentType "application/json" `
            -Body $body

            
        if ($excelUser.Status -eq 'Invited') {
            $excelUser.Status = 'Success'
            Write-Host "Update request sent for user $mail" -ForegroundColor Green
        }

        
    }
    catch {
        $excelUser.Status = 'Failed'
        $excelUser.ErrorStep = 'Update'
        $excelUser.ErrorMessage = $_.Exception.Message
    }


}
$excelUsers | Export-Excel `
    -Path "$($filepath.Replace('.xlsx','_result.xlsx'))" `
    -AutoSize `
    -TableName GuestOnboardingResults


Write-Host "===================================" -ForegroundColor DarkGray
Write-Host "✔ Output file created" -ForegroundColor Green
Write-Host "$($filepath.Replace('.xlsx','_result.xlsx'))" -ForegroundColor Cyan
Write-Host "===================================" -ForegroundColor DarkGray







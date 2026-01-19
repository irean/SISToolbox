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
        [string]$Title  = "Select a file",
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

connect-mggraph -scope "user.invite.all", "user.readwrite.all"

$filepath = Select-FilePath -Title "Select excelfil with user emails" -Filter '*.xlsx'

$users =  Import-Excel -path $filepath

$users | Foreach-Object {
$mail = $_.mail
$givenName = $_.givenName
$surName = $_.surName
$companyName = $_.companyname

$params = @{

    invitedUserEmailAddress = "$mail"
    inviteRedirectUrl       = "https://portal.office.com"
 
}

$guest = Invoke-MggraphRequest -method POST -uri "https://graph.microsoft.com/v1.0/invitations" -ContentType "application/json" -Body $params 



$guest | Foreach-Object {
    $id = $_.invitedUser.ID
    $max = 10
    $delay = 3

    Write-host "$id"

    for ($i = 1; $i -le $max; $i++) {

        try {
            $user = igall "https://graph.microsoft.com/v1.0/users/$id"
                    if ($user) {
            Write-host "User avail after $i attempts"

            $body = @{
                givenName   = "Sandra"
                surName     = "Saluti"
                companyName = "Epical"
            }

            Invoke-MggraphRequest -method PATCH -uri "https://graph.microsoft.com/v1.0/users/$id" -contentType application/json -Body $body
            break
        
        }

        }
        catch {
            Start-Sleep -Seconds $delay 
        }




    }

}
}





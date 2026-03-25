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

# Ensure the Microsoft Graph module is imported
test-Module Microsoft.Graph.Authentication
Write-Host "Disconnecting from previous sessions (if any)" -Foregroundcolor Gray
$connected = Get-MgContext
if ($connected){
    disconnect-Mggraph
}


# Connect to Microsoft Graph
Write-Host "Connecting to tenant, if you dont see login window, check behind others" -ForegroundColor Yellow
Connect-MgGraph -Scopes "Application.Read.All", "Directory.Read.All", "Synchronization.Read.All"  -ContextScope Process

$org = Igall "https://graph.microsoft.com/v1.0/organization"
$orgdisplayName = $org.displayname

Write-Host "You are now connected to $orgdisplayName. Let's proceed!"

function Get-EnterpriseAppsSSOandSCIM {



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
    
$date = Get-Date -Format yyyy-MM-dd
$path = "$folder\$orgdisplayname-AppSSO&ScimReport$date.xlsx"

# Get all Enterprise Applications
Write-Host "Getting all Enterprise Applications" -ForegroundColor Gray
$enterpriseApps = igall  "https://graph.microsoft.com/beta/servicePrincipals" | Where-object {
    $_.servicePrincipalType -like 'Application'
}

$results = @()

$enterpriseApps | Foreach-Object {

    $appname = $_.DisplayName
    $createdDate = $_.CreatedDateTime
    $id = $_.id
    
    Write-Host "Fetching information for app $appname"

    $scimProvisioning = $false

    $provisioningJob = igall "https://graph.microsoft.com/beta/servicePrincipals/$id/synchronization/jobs/"
    if ($provisioningJob) {
        $scimProvisioning = $true
    }
    $ssoMethod = "none"
    $sp = igall "https://graph.microsoft.com/beta/serviceprincipals/$id"
    $prefSigninmethod = $sp.preferredSingleSignOnMode
    # Set value from preferredSingleSignOnMode as SSO
    if ($prefSigninmethod ) {
        $ssomethod = $prefSigninmethod
    }
    # If Non check if logoutURL is set to capture older SSO methods
    elseif ($sp.logouturl) {
        $ssomethod = "Older SAML or OIDC"
    }

    $results += [pscustomobject]@{
        Name             = $appName
        CreatedDate      = $createdDate
        SCIMProvisioning = $scimProvisioning
        SSO              = $ssoMethod
    }
}

$results | Sort-object -Property SSO, SCIMProvisioning -Descending  |  Export-Excel -path $path -TableStyle Medium11 -AutoSize

}

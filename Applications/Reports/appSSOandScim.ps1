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

$connected = Get-MgContext
if ($connected){
    disconnect-Mggraph
}


# Connect to Microsoft Graph
Connect-MgGraph -Scopes "Application.Read.All", "Directory.Read.All", "Synchronization.Read.All"  -ContextScope Process

function Get-EnterpriseAppsSSOandSCIM {

    param (
        [Parameter(mandatory=$true)]
        [String]$path

    )

# Get all Enterprise Applications
$enterpriseApps = igall  "https://graph.microsoft.com/beta/servicePrincipals" | Where-object {
    $_.servicePrincipalType -like 'Application'
}

$results = @()

$enterpriseApps | Foreach-Object {
    $appname = $_.DisplayName
    $createdDate = $_.CreatedDateTime
    $id = $_.id

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

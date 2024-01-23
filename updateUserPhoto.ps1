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

function Update-Photo {
    param (
        [Parameter()]
        [String]$UserPrincipalName,
        [Parameter()]
        [String]$filepath

    )

    Test-module -name Microsoft.Graph.Authentication
    Disconnect-MgGraph
    Connect-MGgraph -scopes user.readwrite.all


    Invoke-MGgraphrequest -method PUT -Uri "https://graph.microsoft.com/beta/users/$userprincipalname/photo/`$value" -ContentType image/jpeg -InputFilePath $filepath

}
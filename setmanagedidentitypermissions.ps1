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

function Set-MIPermissions {
    param (
        [Parameter()]
        [String]$managedIdentityId,
        [Parameter()]
        [String[]]$roles
    )
    
}
test-module -name Microsoft.Graph.Authentication
test-module -name Microsoft.Graph.Applications

Connect-MgGraph -Scopes Application.Read.All, AppRoleAssignment.ReadWrite.All, RoleManagement.ReadWrite.Directory

$Roles | ForEach-Object {

    $roleName = $_

    $mggraph = Get-MgServicePrincipal -Filter "AppId eq '00000003-0000-0000-c000-000000000000'"
    $role = $Mggraph.AppRoles | Where-Object { $_.Value -eq $roleName } 

    New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $managedIdentityId -PrincipalId $managedIdentityId -ResourceId $mggraph.Id -AppRoleId $role.Id

}


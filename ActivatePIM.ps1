function Enable-MGPimRole {

    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [String]$TenantID,
        [Validateset(
            'Application Administrator', 'Exchange Administrator', 'Global Administrator', 'Security Administrator','SharePoint Administrator', 'User Administrator'
        )]
        [Parameter(Mandatory=$true)]
        [String]$role,
        [Parameter(Mandatory=$false)]
        [String]$Justification = 'No Justification Provided'
    )

    #Switching tenants might fail sometimes, disconnecting avoids this
    Disconnect-MgGraph
    #Required module for Connect-MGGraph and Invoke-MgGraphRequest
    import-module Microsoft.Graph.Authentication
    #Make sure to connect with correct scopes and tenant
    Connect-Mggraph -TenantId $TenantID -Scopes RoleManagementPolicy.Read.Directory, RoleManagement.Read.Directory, RoleManagement.Read.All, RoleManagementPolicy.ReadWrite.Directory, RoleManagement.ReadWrite.Directory, User.Read.All
    #To translate a human readable role into a pim request we need to first get the RoleDefinitionID and then we use that to find a policy assignment which we then use to get the policy rules, which contains the PIM activation maximum duration.
    $roleID = Invoke-MGGraphRequest -method GET -Uri https://graph.microsoft.com/beta/roleManagement/directory/roleDefinitions?$top=100 | Select-Object -expandproperty Value | Where-Object {
        $_.displayName -eq $role
    } | Select-Object -ExpandProperty Id 
  
    $policyID = Invoke-MgGraphRequest -Method GET -uri 'https://graph.microsoft.com/beta/policies/roleManagementPolicyAssignments?$filter=scopeId eq %27%2F%27 and scopeType eq %27DirectoryRole%27' | Select-Object -ExpandProperty Value | Where-Object {
        $_.RoleDefinitionId -match $RoleID 
    } | Select-Object -ExpandProperty policyID 
    $mduration = Invoke-MgGraphRequest -method GET -Uri "https://graph.microsoft.com/beta/policies/roleManagementPolicies/$policyID/rules" | Select-Object -Expandproperty Value | Where-Object { $_.id -match 'Expiration_EndUser_Assignment' } | Foreach-Object {
        $_["maximumDuration"]
    }
    #Extracts the id from Userprincipalname from Get-MGcontext
    $u = Get-MgContext | Select-Object -ExpandProperty Account 
    $userID = Get-MGUser -filter "UserPrincipalName eq '$u' " | Select-Object -ExpandProperty Id



    $params = ConvertTo-Json (@{
            action           = "selfActivate"
            principalId      = $userID
            roleDefinitionId = $roleID 
            directoryScopeId = "/"
            justification    = $j
            scheduleInfo     = @{
                startDateTime = Get-Date
                expiration    = @{
                    type     = "AfterDuration"
                    duration = $mduration
                }
            }

        })
    Invoke-MgGraphRequest -Method POST -Uri https://graph.microsoft.com/beta/roleManagement/directory/roleAssignmentScheduleRequests -Body $params
<#
    .SYNOPSIS
        Self Activation of Privileged Role Assignment
    .DESCRIPTION
        Activate a eligible PIM role by suppling tenantID, rolename. Optionaly provide a justification or it defaults to 'No Justification Provided'
        
    .PARAMETER TenantID
        Mandatory value either Tenant domain or guid
            
    .PARAMETER Role
        Mandatory validated values of common built in Roles, add more to list if needed. 
            
    .PARAMETER Justification
        Optional value adding a justification, if not provided a default value 'No Justification Provided' is added. 
    .INPUTS
        None
    .OUTPUTS
        System.String
    .NOTES
        Version:        1.0
        Author:         Sandra Saluti
        Creation Date:  2022-11-13
    .EXAMPLE
        Enable-MGPimRole -TenantID randomcompany.onmicrosoft.com -role 'User Administrator' 
    .EXAMPLE
        Enable-MGPimRole -TenantID a3506fxx-xxxx-4659-8168-xxxxxxxxx -role 'Exchange Administrator' -Justification 'Updating MailFlow' 
#>
    
}







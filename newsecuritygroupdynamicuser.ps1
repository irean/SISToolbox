
function New-MgSecurityGroupDynamicUser {
    param (
        [Parameter()]
        [String]$licname,
        [Parameter()]
        [String]$ServiceplanID,
        [Parameter()]
        [String]$notinclserviceplanID,
        [Parameter()]
        [String[]]$orgnames
    )
    
    $membrulelic = ''
    if ($notinclserviceplanID) {
        $membrulelic = "(user.assignedPlans -any (assignedPlan.servicePlanId -eq `"$serviceplanID`" -and assignedPlan.capabilityStatus -eq `"Enabled`")) -and -not (user.assignedPlans -any (assignedPlan.servicePlanId -eq `"$notinclserviceplanID`" ))"
    }
    else {
        $membrulelic = "(user.assignedPlans -any (assignedPlan.servicePlanId -eq `"$serviceplanID`" -and assignedPlan.capabilityStatus -eq `"Enabled`"))"
    }

    $orgnames | Foreach-Object {
        $orgname = $_
        $o = $orgname -replace ' ', '_'
        $mailnick = "lic_$($o -replace '_','')_$($licname)"
        $membrule = "(user.companyName -eq `"$orgname`") -and  $($membrulelic)"
        $description = "Dynamic User Group based on $($licname) and CompanyName $($orgname) with purpose to count assigned licenses that match CloudIQ subscription"


       
        $params = @{
            description                   = $description
            displayName                   = "lic-audit-$($o)-$($licname)"
            groupTypes                    = @(
                'DynamicMembership'
            )
            mailEnabled                   = $false
            mailnickName                  = $mailnick
            securityEnabled               = $true
            membershipRule                = $membrule 
            membershipRuleProcessingState = 'on'
        }

        New-MgGroup -BodyParameter $params

    }


}
[CmdletBinding()]

# Change values based on your new application and tenant
$DisplayName = "<Name Of Application>"
$description = "<desctription of your application>"
$tenantid = "<tenantID>"

# Create the application
$body = ConvertTo-Json -Depth 20 @{ 
    DisplayName            = "$displayname" 
    description            = "$description"
    signInAudience         = "AzureADMyOrg"
    api                    = @{

    }
    requiredResourceAccess = @(@{
            resourceAppId  = '00000003-0000-0000-c000-000000000000'
            resourceAccess = @(
                @{
                    id   = 'df021288-bdef-4463-88db-98f22de89214'
                    type = 'Role'
                },
                @{
                    id   = '7ab1d382-f21e-4acd-a863-ba3e13f7da61'
                    type = 'Role'
                }                
            )   
        }        
    )
}

$app = Invoke-MgGraphRequest -method POST -ContentType "application/json" -Body $body -uri https://graph.microsoft.com/beta/applications
# Create a client secret
$appid = $app.id 
$clientsecrets = ConvertTo-Json @{
    PasswordCredential = @{
        displayName = "$Displayname Secret"

    }

}
$cientID = $app.appId
$secret = Invoke-MgGraphRequest -Method POST -ContentType "application/json" -Body $clientsecrets -Uri https://graph.microsoft.com/beta/applications/$appid/addPassword

$results = @{
    DisplayName   = $app.displayName
    ApplicationID = $app.appId
    ClientSecret  = $secret.secretText
    tenantID      = $tenantid
}

$results 
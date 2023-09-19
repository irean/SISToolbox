function Connect-MgApplication {
    param (
        [Parameter()]$ApplicationObjectID,
        [Parameter()]$tenantid,
        [Parameter()]$clientSecret
    )

    $appid = $ApplicationObjectID
    $tenantid = $tenantid
    $secret = $clientSecret
 
    $body = @{
        Grant_Type    = "client_credentials"
        Scope         = "https://graph.microsoft.com/.default"
        Client_Id     = $appid
        Client_Secret = $secret
    }

    $connection = Invoke-RestMethod `
        -Uri https://login.microsoftonline.com/$tenantid/oauth2/v2.0/token `
        -Method POST `
        -Body $body
 
    $token = ConvertTo-SecureString -AsPlainText -Force $connection.access_token

    Connect-MgGraph -AccessToken $connection.access_token

    
}
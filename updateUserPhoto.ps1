

function Update-Photo {
    param (
        [Parameter()]
        [String]$UserPrincipalName,
        [Parameter()]
        [String]$filepath

    )

    Disconnect-MgGraph
    Connect-MGgraph -scopes user.readwrite.all


    Invoke-MGgraphrequest -method PUT -Uri "https://graph.microsoft.com/beta/users/$userprincipalname/photo/`$value" -ContentType image/jpeg -InputFilePath $filepath

}
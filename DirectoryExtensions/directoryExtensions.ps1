#Create Directory Extension Azure AD for UserPurpose.

function Create-DirectoryExtensionUserPurpose {
    param (
        [Parameter()]
        [String]$ApplicationObjectID,
        [Parameter()]
        [String]$nameofextension
    )

    $body = @{
        name          = $nameofextension
        datatype      = "String"
        targetObjects = @(
            "User"
        )
    }
    Invoke-MgGraphRequest -method POST "https://graph.microsoft.com/v1.0/applications/$ApplicationObjectID/extensionProperties"  -Body $body
    
}


function List-DirectoryExtension {
    param (
        [Parameter()]
        [String]$AppDisplayname
    )

    if ($AppDisplayname) {
        $app = Invoke-MggraphRequest -Method GET "https://graph.microsoft.com/v1.0/applications?`$Filter=displayname eq '$AppDisplayname'" | Select-Object -ExpandProperty Value | Select-OBject -Expandproperty id 
        Invoke-MgGraphRequest -Method GET "https://graph.microsoft.com/v1.0/applications/$app/extensionProperties" | Select-Object -ExpandProperty value
    }

    else {
        $url = "https://graph.microsoft.com/v1.0/applications"
        do {
            $respons = Invoke-MggraphRequest -method GET -uri $url    
            $url = $respons.'@odata.nextlink'
            $respons.value | Foreach-Object {

                Invoke-MgGraphRequest -Method GET "https://graph.microsoft.com/v1.0/applications/$id/extensionProperties" | Select-Object -ExpandProperty value

            } 
            
        } while (
            $url
        )

        Invoke-MggraphRequest -Method GET "https://graph.microsoft.com/v1.0/applications" | Select-object -ExpandProperty Value |  Foreach-Object {

            $id = $_.id

            Invoke-MgGraphRequest -Method GET "https://graph.microsoft.com/v1.0/applications/$id/extensionProperties" | Select-Object -ExpandProperty value
        }

    }
    
}


function Get-DirectoryExtensionValues {
    param (
        [Parameter()]
        [String]$directoryextensionname,
        [Parameter()]$UserUPN
    )

    if ($UserUPN) {
        Invoke-MggraphRequest -method GET "https://graph.microsoft.com/beta/users/$($UserUPN)?$select=id,displayName,$directoryextensionname" | Select-Object id, DisplayName, $directoryextensionname      
     
    }
    else {

        $url = "https://graph.microsoft.com/beta/users?$select=id,displayName,$directoryextensionname" 

        do {
            $respons = Invoke-MggraphRequest -method GET -uri $url    
            $url = $respons.'@odata.nextlink'
            $respons.value | Foreach-Object {

                Write-Output ($_  | Select-OBject id, displayname, $directoryextensionname, userPrincipalName)

            }
            
        } while (
            $url
        )          

    }
    
}
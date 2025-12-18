function test-module {
    [CmdletBinding()]
    param(
        [String]$Name
  
    )
   Write-Host "Checking module $name"
    if(-not (Get-Module $Name)){
        Write-Host "Module $Name not imported, trying to import"
        try{
            if($Name -eq 'Microsoft.Graph'){
                Write-Host "Microsoft.Graph module import takes a while"
                Import-Module $Name  -ErrorAction Stop
            }
            else{
                Import-Module $Name  -ErrorAction Stop
            }
        
        }
        catch{
            Write-Host "Module $Name not found, trying to install"
            Install-Module $Name -Scope CurrentUser -AllowClobber -Force -AcceptLicense -SkipPublisherCheck
            Write-Host "Importing module  $Name "
            Import-Module $Name  -ErrorAction stop 
        }
    } 
    else{
        Write-Host "Module $Name is imported"
    }   
}
function test-module {
    [CmdletBinding()]
    param(
        [String]$Name
  
    )
    Write-Host "Checking module $name..." -ForegroundColor Cyan
    if (-not (Get-Module $Name)) {
        Write-Host "Module $Name not imported, trying to import..." -ForegroundColor Yellow
        try {
            if ($Name -eq 'Microsoft.Graph') {
                Write-Host "Microsoft.Graph module import takes a while..." -ForegroundColor Yellow
                Import-Module $Name  -ErrorAction Stop
            }
            elseif ($Name -eq 'Az') {
                Write-Host "Module Az is being imported. This might take a while..." -ForegroundColor Yellow
            }
            else {
                Import-Module $Name  -ErrorAction Stop
            }
        
        }
        catch {
            Write-Host "Module $Name not found, trying to install..." -ForegroundColor Yellow
            Install-Module $Name -Scope CurrentUser -AllowClobber -Force -AcceptLicense -SkipPublisherCheck
            Write-Host "Importing module $Name..." -ForegroundColor Yellow
            Import-Module $Name -ErrorAction Stop
        }
    } 
    else {
        Write-Host "Module $Name is already imported." -ForegroundColor Green
    }   
}
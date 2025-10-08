function igall {
  [CmdletBinding()]
  param (
    [string]$Uri,
    [switch]$Eventual,
    [int]$limit=1000
  )
  $nextUri = $uri
  $count=0
  $headers = @{
    Accept = 'application/json'
  }
  if ($Eventual) {
    $headers.Add('ConsistencyLevel', 'eventual')
  }
  do {
    $result = Invoke-MgGraphRequest -Method GET -uri $nextUri -Headers $headers
    $nextUri = $result.'@odata.nextLink'
    if ($result.value) {
      $result.value | ConvertTo-PSCustomObject
    } elseif ($result) {
      $result | ConvertTo-PSCustomObject
    }
    $count +=1
  } while ($nextUri -and ($count -lt $limit))
}

  function ig {
    [CmdletBinding()]
    param (
      [string]$Uri
    )
    $result = Invoke-MgGraphRequest -Method GET -uri $uri
    if ($result.value) {
      $result.value | ConvertTo-PSCustomObject
    } else {
      $result | ConvertTo-PSCustomObject
    }
  }
  function ConvertTo-PSCustomObject {
  [CmdletBinding()]
  param (
    [Parameter(ValueFromPipeline = $true, Mandatory = $true)]
    [System.Collections.Hashtable] $InputObject
  )
  Process {
    if ( $InputObject) {
      $o = New-Object psobject
      foreach ($key in $InputObject.Keys) {
        $value = $InputObject[$key]
        if ($value -and $value.GetType().FullName -match 'System.Object\[\]') {
          if ($value.Count -gt 0 -and $value[0].GetType().FullName -match  'System.Collections.Hashtable') {
            $tempVal = $value | ConvertTo-PSCustomObject
            Add-Member -InputObject $o -NotePropertyName $key -NotePropertyValue $tempVal
          } elseif ($value.Count -gt 0 -and $value[0].GetType().FullName -match  'System.String') {
            $tempVal = $value | ForEach-Object { $_ }
            Add-Member -InputObject $o -NotePropertyName $key -NotePropertyValue $tempVal
          }
        } elseif ($value -and $value.GetType().FullName -match 'System.Collections.Hashtable') {
          Add-Member -InputObject $o -NotePropertyName $key -NotePropertyValue (ConvertTo-PSCustomObject -InputObject $value)
        } else {
          Add-Member -InputObject $o -NotePropertyName $key -NotePropertyValue $value
        }
      }

      Write-Output $o
    }
  }
}

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
# Disconnecting any existing Microsoft Graph session
Write-Host "🔌 Disconnecting previous Microsoft Graph sign-in..." -ForegroundColor Yellow
Disconnect-MgGraph

Write-Host "🔎 Checking module: Microsoft.Graph.Authentication..." -ForegroundColor Blue
Test-Module -Name Microsoft.Graph.Authentication
Write-Host "🔎 Checking module: ImportExcel..." -ForegroundColor Blue
Test-Module -Name ImportExcel 

Write-Host "✅ All required modules are available. Please connect to Microsoft Graph." -ForegroundColor Green


Connect-Mggraph -Scopes "Organization.Read.All", "DeviceManagementManagedDevices.Read.All"

$date = Get-Date -format yyyy-MM-dd 
#Get organization name and remove any spaces in name
$org = (Igall "https://graph.microsoft.com/v1.0/organization" | Select-Object -ExpandProperty displayName) -replace ' ', ''
#save file location
Write-Host "Select folder to save file. If nothing seems to happen, please check if a pop-up window is hidden behind other open windows." -ForegroundColor Red
Add-Type -AssemblyName System.Windows.Forms
$FileBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop') 
}
if ($FileBrowser.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host "No folder selected"
    return
}
$folder = $FileBrowser.SelectedPath

$savepath = "$($folder)\$($org)_Warranty_Bath_Lookup_$($date).xlsx"

#get all devices of manufacturer Lenovo
$devices = IGall "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices" | Where-Object {
    $_.manufacturer -eq 'lenovo'
}
#Foreach devices and add properties

$result = $devices | Foreach-Object {
    $model = ($_.model).Substring(0, 4)

    add-member -InputObject $_ -NotePropertyName 'Machine Type'  -NotePropertyValue $model -Force -PassThru | 
    add-member -NotePropertyName SerialOrIMEI -NotePropertyValue $_.serialNumber -force -PassThru
} | Select-Object 'Machine Type', SerialOrIMEI, comment

#Export to Excel
$result | Export-Excel -path $savepath -WorksheetName Batch 

Write-Host "✅ File saved to: $savepath`nNext step: Visit Lenovo’s warranty site (https://pcsupport.lenovo.com/se/sv/warrantylookup/batchquery) and upload the file to check your warranty details."






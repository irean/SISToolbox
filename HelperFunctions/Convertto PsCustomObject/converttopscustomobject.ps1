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

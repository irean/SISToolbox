# ConvertTo-PSCustomObject

## Overview

`ConvertTo-PSCustomObject` is a small helper function that converts PowerShell
hashtables into structured `PSCustomObject` instances.

It is primarily used to normalize data returned from APIs (for example
Microsoft Graph), where responses often contain deeply nested hashtables that
are awkward to work with in PowerShell.

The function preserves the original structure while making the result easier to
inspect, filter, and export.

---

## Why This Helper Exists

API responses represented as hashtables:

- Do not behave like normal PowerShell objects
- Are difficult to pipe and select from
- Are inconvenient to export to CSV or Excel

This helper solves that by converting the data into proper PowerShell objects
while keeping the structure intact.

---

## Syntax

```powershell
$hashtable | ConvertTo-PSCustomObject
```
or 

```powershell
ConvertTo-PSCustomObject -InputObject $hashtable
```
## Input

### `InputObject`

- **Type:** `System.Collections.Hashtable`
- **Required:** Yes
- **Accepts pipeline input:** Yes

The hashtable to convert into a `PSCustomObject`.  
Nested hashtables and arrays are processed recursively.

---

## Behavior

For each property in the input hashtable, the function:

- Converts nested hashtables into `PSCustomObject`
- Recursively processes arrays of hashtables
- Preserves arrays of strings
- Leaves scalar values unchanged

The output behaves like a native PowerShell object and supports dot-notation,
filtering, and exporting.

---

## Examples

### Simple Hashtable

```powershell
$hash = @{
    Name = "Alice"
    Age  = 30
}

$object = $hash | ConvertTo-PSCustomObject
```

### Nested Hashtable

```powershell
$hash = @{
    User = @{
        DisplayName = "John Doe"
        Enabled     = $true
    }
}

$object = $hash | ConvertTo-PSCustomObject
$object.User.DisplayName
```

### Array of Objects

```powershell
$hash = @{
    Roles = @(
        @{ Name = "Global Administrator"; Enabled = $true }
        @{ Name = "User Administrator"; Enabled = $false }
    )
}

$object = $hash | ConvertTo-PSCustomObject
$object.Roles | Select-Object Name, Enabled
```
### API / Microsoft Graph Example

```powershell
$response = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/organization"
$object   = $response | ConvertTo-PSCustomObject
```
## Output

- **Type:** `PSCustomObject`
- Structure mirrors the original hashtable
- Fully compatible with PowerShell pipelines
- Safe to use for exporting to CSV, Excel, or other reporting formats

---

## When to Use

Use this helper function when:

- Working with Microsoft Graph or other REST API responses
- Handling JSON converted into hashtables
- Preparing data for reporting or export
- You want predictable PowerShell object behavior and dot-notation access

---

## Notes

- Intended as an internal helper function
- Best used early to normalize API responses
- Focused on clarity and reliability rather than performance optimization
- Can handle nested arrays and nested hashtables recursively

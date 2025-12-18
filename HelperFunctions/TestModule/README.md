# test-module

## Overview

`test-module` is a PowerShell helper function designed to **check, import, or install modules** required for your script.  

It helps ensure that the necessary modules are available in the current session, automatically handling missing modules by installing them from the PowerShell Gallery.

This is especially useful in automation scripts where dependencies must be guaranteed before execution.

---

## Syntax

```powershell
test-module -Name <string>
```

## Parameters

| Name | Type   | Required | Description |
|------|--------|----------|-------------|
| `Name` | `string` | Yes | The name of the PowerShell module to check, import, or install. Examples: `Microsoft.Graph`, `Az`, `Az.Accounts`, `ImportExcel`. |

---

## Behavior

- Checks if the specified module is already imported in the current session.
- If the module is **not imported**, it attempts to import it.
- If the module **cannot be imported**, it automatically installs it from the PowerShell Gallery and then imports it.
- Provides user-friendly output messages with color:
  - **Cyan** for checking steps
  - **Yellow** for warnings or installation steps
  - **Green** for success messages
- Handles specific modules like `Microsoft.Graph` and `Az` with tailored messages, as they may take longer to import.

---

## Examples

```powershell
# Check and import the Microsoft.Graph module
test-module -Name "Microsoft.Graph"

# Check and import the Az module
test-module -Name "Az"

# Check and import a custom module
test-module -Name "ImportExcel"
```

## Notes

- The function ensures that the current session has access to the requested module, preventing errors due to missing dependencies.
- Modules are installed **for the current user only** using `Install-Module -Scope CurrentUser`.
- `-AllowClobber` and `-Force` are used during installation to avoid conflicts and automatically accept license prompts.
- Especially useful in automation scripts or when onboarding new environments where modules may not be preinstalled.
- Provides informative, color-coded output to make it clear what is happening:
  - **Cyan**: Checking steps
  - **Yellow**: Warnings or installation steps
  - **Green**: Success messages

---

## Related Functions

- `Connect-MgGraph` — Ensures connection to Microsoft Graph before running queries.
- `igall` — Fetches paginated results from Microsoft Graph.
- `ConvertTo-PSCustomObject` — Converts hashtables to structured PS objects for easier processing.

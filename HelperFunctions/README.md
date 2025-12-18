# PowerShell Helper Functions

## Overview

This folder contains a collection of **PowerShell helper functions I have created** to simplify common tasks, improve automation, and reduce repetitive code in my scripts.  

These functions are designed to be **reusable across multiple scripts** and often rely on each other to provide consistent behavior.  


---

## Purpose

- Consolidate reusable PowerShell logic I have developed over time.
- Make scripts cleaner, easier to maintain, and faster to write.
- Handle repetitive operations automatically, such as module checks, Graph API requests, or data conversion.
- Provide small, focused functions that can be combined for larger automation tasks.

---

## Structure

Each helper function has its own folder or script file, ideally with a dedicated `README.md` describing:

- Overview of the function  
- Parameters  
- Behavior and important notes  
- Examples  
- Related functions  

---

## Current Functions

| Function | Description |
|----------|-------------|
| `ConvertTo-PSCustomObject` | Converts nested hashtables into `PSCustomObject` for easier dot-notation access and pipeline use. |
| `igall` | Fetches paginated results from Microsoft Graph and converts them to `PSCustomObject`. Handles eventual consistency and large datasets. |
| `ig` | Fetches single-page results from Microsoft Graph and converts them to `PSCustomObject`. |
| `test-module` | Checks if a PowerShell module is installed, imports it, or installs it automatically from the gallery. |

> Each function has its own README.md with usage instructions, parameters, and examples.

---

## Usage Guidelines

1. Dot-source or import the helper functions in your script.
2. Ensure required dependencies (e.g., `Connect-MgGraph`) are loaded first.
3. Use helpers individually or combine them to simplify complex scripts.
4. Refer to each function’s README for detailed guidance.

---

## Future Additions

- New helpers will be added as I create them.
- Functions may evolve; breaking changes will be documented in the function-specific README.
- Goal: Maintain a **reusable, personal toolbox** for consistent scripting.

---

## Example Usage

```powershell
# Example: using helper functions in your script
. "C:\Path\To\HelperFunctions\ConvertTo-PSCustomObject.ps1"
. "C:\Path\To\HelperFunctions\igall.ps1"
. "C:\Path\To\HelperFunctions\test-module.ps1"

# Now you can call them in your scripts
test-module -Name "Microsoft.Graph.Authentication"
$users = igall -Uri "https://graph.microsoft.com/v1.0/users"
```



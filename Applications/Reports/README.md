# Enterprise Apps SSO and SCIM Report Script

This Powershell script collects information about **Enterprise Applications** in Entra ID using Microsoft Graph API. It provides details on the Single Sign-On (SSO) method and SCIM provisioning status of each application and exports the results to an Excel file.

## Prerequisites
1. **Powershell 7**
2. **Microsoft Graph Module**:
   - The script will automatically install and import required modules (`Microsoft.Graph.Authentication`, `ImportExcel`) if not ready available.
3. **Permissions**:
   - the script will connect to Microsoft graph with scopes `Application.Read.All`, `Directory.Read.All`, `Synchronization.Read.All`

## Logic

- Retrieves all Enterprise Applications.
- Checks for SSO method and SCIM provisioning status.
- Checks for logoutURL to capture older methods.
- Exports the results into an Excel file.

## Example
```powershell
Get-EnterpriseAppsSSOandSCIM -path "C:\path\to\output.xlsx"
```



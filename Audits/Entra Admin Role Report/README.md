# Entra ID Admin & Role Report

PowerShell script that generates an Excel report of:

- **Active Entra ID administrators**
- **PIM eligible roles**
- **Azure RBAC role assignments**
- **User authentication methods and MFA status**

The report helps identify privileged users and review MFA posture for administrative accounts.

---

## Features

- Retrieves **Entra ID directory role members**
- Retrieves **PIM eligible role assignments**
- Expands **group-based role assignments**
- Collects **user sign-in activity**
- Reports **registered authentication methods**
- Detects **strong MFA methods**
- Retrieves **Azure RBAC role assignments across subscriptions**
- Exports results to **Excel**

---

## Requirements

PowerShell **7+**

Required modules (installed automatically if missing):

- Az.Accounts
- Az.Resources
- Microsoft.Graph.Authentication
- ImportExcel

### Required Graph permissions:

- `RoleManagement.Read.Directory`
- `User.Read.All`
- `GroupMember.Read.All`
- `Directory.Read.All`
- `RoleEligibilitySchedule.Read.Directory`
- `RoleManagement.Read.All`
- `AuditLog.Read.All`


Recommended roles:

- **Global Reader**
- **Azure Reader**

---

## Usage

Run the script:

```powershell
.\Generate-EntraAdminReport.ps1
```

### When running the script you will be prompted to:

1. Sign in to Azure  
2. Sign in to Microsoft Graph  
3. Select a folder where the report will be saved



---

## Output

The script generates an Excel file:

```powershell
<OrganizationName>-EntraIDAdminReport-YYYY-MM-DD.xlsx
```
### Worksheets

| Worksheet | Description |
|-----------|-------------|
| Administrators | Active Entra ID directory role members |
| Eligible Roles | Users eligible for roles through Privileged Identity Management (PIM) |
| Azure Roles | Azure RBAC role assignments across subscriptions |

---

## Authentication Information

For each user the script collects:

- Last sign-in time
- Last password change
- Registered authentication methods
- Strong MFA status

Strong authentication methods detected:

- Microsoft Authenticator
- Passwordless Microsoft Authenticator
- FIDO2 security keys

---

## Notes

- The script uses **Microsoft Graph beta endpoints**.
- Large tenants may take longer to process due to **Graph pagination and group expansion**.

## Example Report

An example report is available in the repository:



The report contains three worksheets.

### Administrators

Lists users that currently have **active Entra ID directory roles**.

Includes additional security insights such as:

- Last sign-in time
- Last password change
- Registered authentication methods
- Strong MFA detection

Example columns:

| Column | Description |
|------|------|
| Role | Assigned directory role |
| DisplayName | User name |
| UserPrincipalName | Sign-in name |
| CompanyName | Company attribute |
| AccountEnabled | Indicates if the account is enabled |
| CreatedDateTime | Account creation time |
| LastPasswordChangeDateTime | Last password change |
| LastSignInDateTime | Most recent sign-in |
| hasStrongMFA | Indicates presence of strong MFA |
| StrongAuthCount | Number of strong authentication methods |

---

### Eligible Roles

Lists users who are **eligible for privileged roles via PIM**.

Both **direct assignments** and **group-based assignments** are included.

Example columns:

| Column | Description |
|------|------|
| DisplayName | User name |
| UserPrincipalName | Sign-in name |
| EligibleRole | Role available via PIM |
| MemberType | Direct or group assignment |
| EligibleRoleGroup | Group providing eligibility |

---

### Azure Roles

Lists **Azure RBAC role assignments across subscriptions**.

Includes:

- Users
- Groups
- Service principals

Example columns:

| Column | Description |
|------|------|
| RoleDefinitionName | Azure RBAC role |
| DisplayName | Identity name |
| SigninName | User sign-in |
| ObjectId | Azure AD object ID |
| ObjectType | Type of identity |
| Subscription | Subscription scope admiadm
# igall and ig

## Overview

These are helper functions for fetching data from Microsoft Graph in PowerShell.  

- **`igall`**: Handles paginated responses and can fetch large datasets, optionally using eventual consistency.  
- **`ig`**: Fetches a single page of data without pagination.

Both functions convert the raw JSON/hashtable responses from Microsoft Graph into structured `PSCustomObject` instances using `ConvertTo-PSCustomObject`, making it easier to work with in PowerShell pipelines.

---

## igall

### Purpose

- Fetches all results from a Microsoft Graph endpoint, handling pagination automatically.
- Converts responses into `PSCustomObject`.
- Supports eventual consistency headers for certain queries.

### Syntax

```powershell
igall -Uri <string> [-Eventual] [-limit <int>]
```

## Parameters

### igall

| Name      | Type     | Required | Description |
|-----------|----------|----------|-------------|
| `Uri`     | `string` | Yes      | The Microsoft Graph endpoint to query. |
| `Eventual`| `switch` | No       | Adds the `ConsistencyLevel: eventual` header, used for queries that require eventual consistency. |
| `limit`   | `int`    | No       | Maximum number of pages to fetch. Default is `1000`. |

### ig

| Name  | Type     | Required | Description |
|-------|----------|----------|-------------|
| `Uri` | `string` | Yes      | The Microsoft Graph endpoint to query. |

---

## Behavior

### igall

- Starts at the provided `Uri` and fetches data using `Invoke-MgGraphRequest`.
- Handles paginated responses automatically using `@odata.nextLink`.
- Continues fetching until all pages are retrieved or the specified `limit` is reached.
- Converts each result into a `PSCustomObject` for easy PowerShell usage.
- Supports eventual consistency by adding the `ConsistencyLevel: eventual` header when `-Eventual` is used.
- Handles arrays of objects and single objects correctly.

### ig

- Sends a GET request to the specified URI using `Invoke-MgGraphRequest`.
- If the response contains a `value` array, it converts the array into `PSCustomObject`.
- If the response is a single object, it is converted into `PSCustomObject`.
- Does not handle pagination; ideal for single-page queries.

---

## Examples

### igall Example

```powershell
# Fetch all users from Microsoft Graph
$users = igall -Uri "https://graph.microsoft.com/v1.0/users"

# Fetch all groups with eventual consistency
$groups = igall -Uri "https://graph.microsoft.com/v1.0/groups" -Eventual
```
### ig Example

```powershell
# Fetch a single page of users
$usersPage = ig -Uri "https://graph.microsoft.com/v1.0/users?$top=50"
```
## Notes

- Both functions require an active Microsoft Graph connection (`Connect-MgGraph`) with the appropriate permissions.
- `igall` is intended for large datasets and automatically handles pagination.
- `ig` is designed for single-page queries and is faster for smaller datasets.
- All returned objects are processed by `ConvertTo-PSCustomObject`, allowing easy dot-notation access and pipeline operations.
You can find `ConvertTo-PSCustomObject` in the repository: [ConvertTo-PSCustomObject Helper](https://github.com/irean/SISToolbox/tree/main/HelperFunctions/Convertto%20PsCustomObject)
- Use `-Eventual` with `igall` for queries that need eventual consistency, such as certain `$count` queries or filtered searches.
- The `$limit` parameter in `igall` prevents infinite loops when paging through very large datasets.

---

## Related Commands

- `Invoke-MgGraphRequest` — Performs the raw HTTP request to Microsoft Graph.
- `ConvertTo-PSCustomObject` — Converts hashtables and arrays into structured PowerShell objects.
- `Connect-MgGraph` — Establishes a Microsoft Graph session with proper authentication.


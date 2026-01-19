# Bulk Guest Invitation & Update Script (Microsoft Graph)

This PowerShell script bulk-invites guest users to an Entra ID tenant and updates user attributes based on data from an Excel file.

The script is designed to be safe, auditable, and resilient to Microsoft Graph edge cases such as eventual consistency and partial directory objects.

The script never stops on a single-user failure. Instead, it records the outcome for each user and exports the results to a new Excel file.

---

## Features

* Bulk invite guest users via Microsoft Graph
* Update guest or existing user attributes after creation
* Handles eventual consistency with retry logic
* Detects existing users and avoids duplicate invitations
* Writes per-user status and errors back to Excel
* Produces an auditable result file
* Designed for safe re-runs

---

## Prerequisites

### PowerShell

* PowerShell 5.1 or PowerShell 7+

### Required Modules

The script will automatically install missing modules if needed:

* Microsoft.Graph.Authentication
* ImportExcel

### Microsoft Graph Permissions

The account used must have consent for:

* User.Invite.All
* User.ReadWrite.All
* Organization.Read.All

Admin consent is required.

---

## Input Excel File

The input Excel file must contain one user per row.

### Required column

| Column | Description                              |
| ------ | ---------------------------------------- |
| mail   | External email address of the guest user |

### Optional columns

Any additional columns will be treated as Microsoft Graph **user attributes** and sent in the PATCH request.

This allows you to manage and update standard Entra ID / Microsoft Graph user properties directly from Excel.

Examples:

* givenName
* surname
* companyName
* city
* department
* jobTitle

The full list of supported user properties is documented in Microsoft Graph:

[https://learn.microsoft.com/graph/api/resources/user](https://learn.microsoft.com/graph/api/resources/user)

Column names must match Microsoft Graph user property names exactly.

---

## Script Flow

1. Connect to Microsoft Graph
2. Confirm the connected tenant
3. Import users from Excel
4. For each user:

   * Build an update body from Excel attributes
   * Check if a usable user already exists
   * Invite the user if needed
   * Wait for user availability (retry logic)
   * Update user attributes
   * Record status and errors
5. Export results to a new Excel file

---

## User Existence Logic

A user is considered existing only if Microsoft Graph returns a user object with a valid id.

This avoids false positives caused by:

* Shadow objects
* Soft-deleted guests
* Directory artifacts that contain mail but no usable identity

If no usable ID is found, the script will invite the user instead.

---

## Retry Logic

Newly invited guests may not be immediately available in the directory.

The script:

* Retries lookup up to 10 times
* Waits 3 seconds between attempts
* Only retries newly invited users

If the user never becomes available, the row is marked as Skipped.

---

## Result and Error Model

Each input row is enriched with additional columns and exported to a new Excel file.

### Output columns

| Column       | Meaning                                                        |
| ------------ | -------------------------------------------------------------- |
| Status       | Processing result (Success, Invited, Updated, Skipped, Failed) |
| UserId       | Graph user ID if resolved                                      |
| ErrorStep    | Step where the error occurred (Lookup, Invite, Retry, Update)  |
| ErrorMessage | Human-readable error message                                   |

### Status meanings

| Status  | Description                                 |
| ------- | ------------------------------------------- |
| Success | User invited and updated successfully       |
| Invited | User invited, update not yet completed      |
| Updated | Existing user updated                       |
| Skipped | User could not be resolved after invitation |
| Failed  | Graph API error occurred                    |

---

## Output File

After execution, a new Excel file is created:

<original_filename>_result.xlsx

The output file contains:

* All original input columns
* Status and error columns
* A full audit trail of the run

---

## Error Handling Philosophy

* Errors are treated as data, not fatal conditions
* The script never stops on a single-user failure
* All failures are recorded per row
* Write-Host is used for runtime visibility only

---

## Safe Re-runs

The script is designed to be rerun safely:

* Existing users are detected and updated
* Failed or skipped users can be filtered and retried
* No duplicate invitations are sent when a valid user exists

---

## Common Scenarios

| Scenario                          | Outcome                   |
| --------------------------------- | ------------------------- |
| User already exists               | Attributes updated        |
| User never existed                | Invited and updated       |
| User invited but not materialized | Skipped                   |
| PATCH fails                       | Failed with error message |

---

## Notes and Limitations

* Filtering on mail can return non-user objects; the script guards against this
* Guest UPNs differ from external email addresses
* Some Graph responses are eventually consistent

---

## Intended Audience

This script is intended for:

* Identity and Access Management engineers
* Entra ID administrators
* Automation and governance teams

It assumes familiarity with:

* PowerShell
* Microsoft Graph
* Entra ID concepts

---

## License and Usage

Internal use script. Adapt and extend as needed for your organization.

---

## Summary

This script provides a robust, production-ready approach to bulk guest onboarding with Microsoft Graph.

It is designed to reflect real-world Graph behavior rather than idealized assumptions.

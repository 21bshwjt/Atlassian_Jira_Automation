# Atlassian Jira Automation

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%20%7C%207%2B-5391FE?logo=powershell&logoColor=white)
![Jira Cloud](https://img.shields.io/badge/Jira%20Cloud-REST%20API%20v3-0052CC?logo=jira&logoColor=white)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20Linux%20%7C%20macOS-lightgrey)

A collection of PowerShell scripts for automating common **Jira Cloud** tasks: validating API access, bulk-copying rich-text content between issues, and exporting project issues to formatted Excel reports.

All scripts talk directly to the [Jira Cloud REST API v3](https://developer.atlassian.com/cloud/jira/platform/rest/v3/) using email + API token authentication. No Jira plugins or admin rights are required.

---

## Table of Contents

- [Scripts at a Glance](#scripts-at-a-glance)
- [Prerequisites](#prerequisites)
- [Getting Started](#getting-started)
- [Usage](#usage)
  - [Test-JiraApiToken.ps1](#1-test-jiraapitokenps1)
  - [Copy-JiraDescriptionAndAC.ps1](#2-copy-jiradescriptionandacps1)
  - [Export-JiraAllIssues.ps1](#3-export-jiraallissuesps1)
  - [Export-JiraIssuesByLabel.ps1](#4-export-jiraissuesbylabelps1)
- [Custom Field IDs](#custom-field-ids)
- [Excel Report Formatting](#excel-report-formatting)
- [Security Best Practices](#security-best-practices)
- [Troubleshooting](#troubleshooting)
- [Project Structure](#project-structure)

---

## Scripts at a Glance

| Script | Purpose | Output |
|---|---|---|
| [`Test-JiraApiToken.ps1`](Scripts/Test-JiraApiToken.ps1) | Verifies that your email + API token can reach the Jira REST API, with detailed diagnostics on failure | Console + result object |
| [`Copy-JiraDescriptionAndAC.ps1`](Scripts/Copy-JiraDescriptionAndAC.ps1) | Copies **Description** and **Acceptance Criteria** from one source issue to one or many target issues, preserving full ADF rich-text formatting | Updated Jira issues |
| [`Export-JiraAllIssues.ps1`](Scripts/Export-JiraAllIssues.ps1) | Exports every issue in a project to a single Excel worksheet | `.xlsx` workbook |
| [`Export-JiraIssuesByLabel.ps1`](Scripts/Export-JiraIssuesByLabel.ps1) | Exports issues grouped by label, one styled worksheet per label | `.xlsx` workbook |

---

## Prerequisites

| Requirement | Notes |
|---|---|
| **PowerShell** | Windows PowerShell 5.1 or PowerShell 7+ |
| **Jira Cloud account** | With browse/edit permission on the target project(s) |
| **Atlassian API token** | Create one at [id.atlassian.com → Security → API tokens](https://id.atlassian.com/manage-profile/security/api-tokens) |
| **ImportExcel module** | Required by the export scripts only. Installed automatically for the current user if missing |

To install `ImportExcel` manually:

```powershell
Install-Module -Name ImportExcel -Scope CurrentUser
```

---

## Getting Started

```powershell
# 1. Clone the repository
git clone <repository-url>
cd Atlassian_Jira_Automation\Scripts

# 2. (Windows) Allow local scripts to run for this session
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

# 3. Store your credentials in variables (avoids leaking them into history)
$Email    = 'you@company.com'
$ApiToken = Read-Host 'Jira API token'

# 4. Confirm API access before running anything else
. .\Test-JiraApiToken.ps1
Test-JiraApiAccess -Email $Email -ApiToken $ApiToken -BaseUrl 'https://your-site.atlassian.net'
```

> **Tip:** Every script accepts a `-JiraBaseUrl` (or `-BaseUrl`) parameter. Always pass your own site URL rather than relying on the built-in defaults.

---

## Usage

### 1. `Test-JiraApiToken.ps1`

Defines the `Test-JiraApiAccess` function, which calls `/rest/api/3/myself` to confirm your credentials work. On failure it reports the HTTP status code, the `X-Authentication-Denied-Reason` and `X-Aip-Blocked` headers (useful for spotting IP allowlist restrictions), and the raw response body.

```powershell
. .\Test-JiraApiToken.ps1
Test-JiraApiAccess -Email $Email -ApiToken $ApiToken -BaseUrl 'https://your-site.atlassian.net' -Verbose
```

| Parameter | Required | Description |
|---|:---:|---|
| `Email` | ✅ | Atlassian account email |
| `ApiToken` | ✅ | Atlassian API token |
| `BaseUrl` | | Jira Cloud site URL (must start with `https://`) |

**Returns:** a `[pscustomobject]` with `Success`, `StatusCode`, `Uri`, `Email`, and, on failure, `DeniedReason`, `IpBlocked`, and `Body`.

---

### 2. `Copy-JiraDescriptionAndAC.ps1`

Reads the **Description** and **Acceptance Criteria** fields from a source issue and writes them to one or more target issues. Content is copied as raw [Atlassian Document Format (ADF)](https://developer.atlassian.com/cloud/jira/platform/apis/document/structure/), so tables, task lists, panels, and nested lists come through unchanged.

```powershell
# Single target
.\Copy-JiraDescriptionAndAC.ps1 -Email $Email -ApiToken $ApiToken `
    -JiraBaseUrl 'https://your-site.atlassian.net' `
    -SourceIssueKey 'PROJ-100' -TargetIssueKeys 'PROJ-101'

# Multiple targets (array or comma-separated string both work)
.\Copy-JiraDescriptionAndAC.ps1 -Email $Email -ApiToken $ApiToken `
    -JiraBaseUrl 'https://your-site.atlassian.net' `
    -SourceIssueKey 'PROJ-100' -TargetIssueKeys 'PROJ-101, PROJ-102, PROJ-103'
```

| Parameter | Required | Description |
|---|:---:|---|
| `Email` | ✅ | Atlassian account email |
| `ApiToken` | ✅ | Atlassian API token |
| `TargetIssueKeys` | ✅ | One or more target keys. Accepts arrays, comma-separated strings, or pipeline input |
| `JiraBaseUrl` | | Jira Cloud site URL |
| `SourceIssueKey` | | Issue to copy content from |
| `AcceptanceCriteriaFieldId` | | Custom field ID for Acceptance Criteria (default `customfield_10320`) |

> ⚠️ **This script overwrites** the Description and Acceptance Criteria on every target issue. Test it on a single issue first. Jira's issue history keeps the previous values if you need to restore them.

---

### 3. `Export-JiraAllIssues.ps1`

Fetches **every issue** in a project (ordered by key) and writes them to a single worksheet, including labels and content-quality flags.

```powershell
.\Export-JiraAllIssues.ps1 -Email $Email -ApiToken $ApiToken `
    -JiraBaseUrl 'https://your-site.atlassian.net' `
    -ProjectKey 'PROJ' -OutputFolder 'D:\Reports'
```

| Parameter | Required | Description |
|---|:---:|---|
| `Email` | ✅ | Atlassian account email |
| `ApiToken` | ✅ | Atlassian API token |
| `JiraBaseUrl` | | Jira Cloud site URL |
| `ProjectKey` | | Jira project key to export |
| `OutputFolder` | | Folder for the workbook (created if missing) |
| `ExcelFileName` | | Output file name (default `JIRA_Identity_Report_All.xlsx`) |
| `WorksheetName` | | Worksheet name (default `All_Issues`) |
| `StartDateFieldId` | | Custom field ID for Start Date (default `customfield_10015`) |
| `AcceptanceCriteriaFieldId` | | Custom field ID for Acceptance Criteria (default `customfield_10320`) |

**Columns:** `Key` · `Summary` · `IssueType` · `Status` · `Assignee` · `Labels` · `StartDate` · `DueDate` · `Description` · `AcceptanceCriteria`

---

### 4. `Export-JiraIssuesByLabel.ps1`

Runs one JQL query per label and writes each result set to its own worksheet with a distinct table colour. Handy for team- or workstream-level reporting.

```powershell
.\Export-JiraIssuesByLabel.ps1 -Email $Email -ApiToken $ApiToken `
    -JiraBaseUrl 'https://your-site.atlassian.net' `
    -ProjectKey 'PROJ' -OutputFolder 'C:\Reports'
```

Parameters match `Export-JiraAllIssues.ps1`, except there is no `WorksheetName` (sheets are named per label). The default file name is `JIRA_Identity_Report.xlsx`.

**Default label → worksheet mapping:**

| Label | Worksheet | Table style |
|---|---|---|
| `entraid` | EntraID | Medium2 (blue) |
| `saviynt` | Saviynt | Medium6 (green) |
| `pki` | PKI | Medium9 (orange) |
| `cyberark` | CyberArk | Medium15 (purple) |

To report on different labels, edit the `$labelMap` hashtable near the top of the script:

```powershell
$labelMap = [ordered]@{
    'backend'  = @{ Sheet = 'Backend';  Style = 'Medium2' }
    'frontend' = @{ Sheet = 'Frontend'; Style = 'Medium6' }
}
```

---

## Custom Field IDs

Custom field IDs differ between Jira sites. The defaults in these scripts are:

| Field | Default ID |
|---|---|
| Start Date | `customfield_10015` |
| Acceptance Criteria | `customfield_10320` |

To find the correct IDs for your site:

```powershell
$auth = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("${Email}:${ApiToken}"))
Invoke-RestMethod -Uri 'https://your-site.atlassian.net/rest/api/3/field' `
    -Headers @{ Authorization = "Basic $auth" } |
    Where-Object name -match 'Acceptance|Start date' |
    Select-Object id, name
```

Then pass them with `-AcceptanceCriteriaFieldId` and `-StartDateFieldId`.

---

## Excel Report Formatting

Both export scripts produce report-ready workbooks with:

- **Auto-filter** on every column and a **frozen header row**
- **Styled Excel tables**
- **Conditional formatting** on Status and on the content-quality flags:

| Value | Colour |
|---|---|
| `Done` | 🟩 Green |
| `In Progress` | 🟨 Yellow |
| `Blocked` | 🟥 Red |
| `To Do` / `Backlog` | ⬜ Light grey |
| `On Hold` | 🟦 Steel blue |
| `TRUE` (field is populated) | Soft green |
| `FALSE` (field is empty) | Soft red |

The `Description` and `AcceptanceCriteria` columns are **booleans**. They show whether each issue has that field filled in, which makes it easy to spot under-specified tickets.

If a query returns no issues, a header-only sheet is still written so the workbook layout stays consistent. An existing workbook at the output path is **replaced** on each run.

---

## Security Best Practices

- **Never commit API tokens.** Pass them at runtime, or load them from a secret store such as [`Microsoft.PowerShell.SecretManagement`](https://learn.microsoft.com/powershell/utility-modules/secretmanagement/overview):
  ```powershell
  $ApiToken = Get-Secret -Name JiraApiToken -AsPlainText
  ```
- Use `Read-Host` (as shown above) so the token never lands in your shell history.
- API tokens carry **your** permissions. Scope them to what you need and revoke unused ones at [id.atlassian.com](https://id.atlassian.com/manage-profile/security/api-tokens).
- Add report output folders to `.gitignore` if they sit inside the repo, because exported data may be confidential.

---

## Troubleshooting

| Symptom | Likely cause and fix |
|---|---|
| `401 Unauthorized` | Wrong email or token, or the token was revoked. Regenerate it and re-test with `Test-JiraApiAccess` |
| `403 Forbidden` with an `X-Aip-Blocked` header | Your organisation enforces an IP allowlist. Connect via VPN or ask an admin to allowlist your IP |
| `404 Not Found` on an issue | The issue key is wrong, or you lack browse permission on the project |
| `400 Bad Request` on update | The field ID is wrong, or the field isn't on the target issue's edit screen |
| Acceptance Criteria column is always `FALSE` | The custom field ID doesn't match your site. See [Custom Field IDs](#custom-field-ids) |
| `cannot be loaded because running scripts is disabled` | Run `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` |
| `Install-Module` fails | Install `ImportExcel` manually, or trust PSGallery: `Set-PSRepository PSGallery -InstallationPolicy Trusted` |

---

## Project Structure

```
Atlassian_Jira_Automation/
├── README.md
└── Scripts/
    ├── Test-JiraApiToken.ps1           # Validate API credentials and diagnose access issues
    ├── Copy-JiraDescriptionAndAC.ps1   # Bulk-copy Description + Acceptance Criteria (ADF)
    ├── Export-JiraAllIssues.ps1        # Export all project issues to Excel
    └── Export-JiraIssuesByLabel.ps1    # Export issues by label, one sheet per label
```

---

## Author

**Biswajit Biswas**

Contributions, issues, and feature requests are welcome.

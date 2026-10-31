<#
.SYNOPSIS
    Pulls Jira label-based queries (one per entry in -Labels) and exports them
    into a single Excel workbook, one worksheet per label, each with its own
    header color, using the ImportExcel module.
   
    Includes boolean (True/False) presence flags for Description & Acceptance Criteria
    with automated color-coded conditional formatting (Green for True, Red for False).

.NOTES
    - Uses Jira Cloud REST API v3 (/rest/api/3/search/jql) with token-based pagination.
    - Requires ImportExcel module (auto-installed if missing).

.EXAMPLE
    .\Export-JiraIssuesByLabel.ps1 -Email me@x.com -ApiToken $tok -Labels label1,label2
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$JiraBaseUrl = "https://your-site.atlassian.net",

    [Parameter(Mandatory = $true, HelpMessage = "Your Atlassian account email")]
    [string]$Email,

    [Parameter(Mandatory = $true, HelpMessage = "Your Atlassian API Token")]
    [string]$ApiToken,

    # One worksheet is created per label, named after the label
    [Parameter(Mandatory = $true, HelpMessage = "Labels")]
    [string[]]$Labels,

    [Parameter(Mandatory = $false)]
    [string]$ProjectKey = "KAN",

    [Parameter(Mandatory = $false)]
    [string]$OutputFolder = "C:\Reports",

    [Parameter(Mandatory = $false)]
    [string]$ExcelFileName = "JIRA_Identity_Report.xlsx",

    # "Start date" field ID (name collides with other fields, so use ID directly)
    [Parameter(Mandatory = $false)]
    [string]$StartDateFieldId = "customfield_10015",

    # "Acceptance criteria" field ID
    [Parameter(Mandatory = $false)]
    [string]$AcceptanceCriteriaFieldId = "customfield_10320"
)

# ---- 1. Setup & Authentication ---------------------------------------------

# Prepare Basic Auth Headers
$authBytes = [System.Text.Encoding]::ASCII.GetBytes("${Email}:${ApiToken}")
$base64Auth = [System.Convert]::ToBase64String($authBytes)

$headers = @{
    "Authorization" = "Basic $base64Auth"
    "Content-Type"  = "application/json"
    "Accept"        = "application/json"
}

# Ensure ImportExcel module is available
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "ImportExcel module not found. Installing for current user..." -ForegroundColor Yellow
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}
Import-Module ImportExcel

$excelPath = Join-Path $OutputFolder $ExcelFileName
New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null

# Remove any previous copy so Export-Excel starts clean (avoids stacking old sheets)
if (Test-Path $excelPath) {
    Remove-Item $excelPath -Force
}

# Helper function to evaluate whether an ADF field or string is populated
function Test-HasContent ($fieldValue) {
    if ($null -eq $fieldValue) { return $false }
    # ADF object check: verify content array exists and has nodes
    if ($fieldValue.PSObject.Properties['content']) {
        return ($fieldValue.content.Count -gt 0)
    }
    # Plain string check
    if ($fieldValue -is [string]) {
        return (-not [string]::IsNullOrWhiteSpace($fieldValue))
    }
    return $true
}

# ---- 2. REST API v3 JQL Search (Token-based pagination) --------------------

function Get-JiraIssuesByJql {
    param(
        [Parameter(Mandatory = $true)][string]$Jql,
        [Parameter(Mandatory = $true)][string[]]$Fields
    )

    $server = $JiraBaseUrl.TrimEnd('/')
    $fieldList = $Fields -join ','
    $token = $null

    do {
        $uri = "{0}/rest/api/3/search/jql?jql={1}&fields={2}&maxResults=100" -f `
            $server, `
            [uri]::EscapeDataString($Jql), `
            [uri]::EscapeDataString($fieldList)

        if ($token) {
            $uri += "&nextPageToken=$([uri]::EscapeDataString($token))"
        }

        try {
            $resp = Invoke-RestMethod -Method Get -Uri $uri -Headers $headers -ErrorAction Stop
            $resp.issues
            $token = $resp.nextPageToken
        }
        catch {
            Write-Error "Failed fetching issues for JQL [$Jql]: $($_.Exception.Message)"
            if ($_.Exception.Response) {
                $stream = $_.Exception.Response.GetResponseStream()
                $reader = [System.IO.StreamReader]::new($stream)
                Write-Error "API Response: $($reader.ReadToEnd())"
            }
            break
        }
    } while ($token -and -not $resp.isLast)
}

# ---- 3. Label -> Worksheet Mapping & Table Styles --------------------------
# Labels come from -Labels; table styles rotate through this palette.

$styles = @(
    'Medium2'    # blue
    'Medium6'    # green
    'Medium9'    # orange/gold
    'Medium15'   # purple
    'Medium3'    # orange
    'Medium7'    # aqua
)

# Accept "a,b,c" as a single argument as well as an array
$Labels = @($Labels -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -Unique)

# ---- 4. Conditional Formatting Rules ---------------------------------------

$ConditionalText = @(
    # --- Status column rules ---
    New-ConditionalText -Text 'Done'        -ConditionalTextColor White -BackgroundColor Green
    New-ConditionalText -Text 'In Progress' -ConditionalTextColor Black -BackgroundColor Yellow
    New-ConditionalText -Text 'Blocked'     -ConditionalTextColor White -BackgroundColor Red
    New-ConditionalText -Text 'To Do'       -ConditionalTextColor Black -BackgroundColor LightGray
    New-ConditionalText -Text 'Backlog'     -ConditionalTextColor Black -BackgroundColor LightGray
    New-ConditionalText -Text 'On Hold'     -ConditionalTextColor White -BackgroundColor SteelBlue

    # --- Description & Acceptance Criteria boolean rules (soft green / soft red) ---
    New-ConditionalText -Text 'TRUE'        -ConditionalTextColor DarkGreen -BackgroundColor LightGreen
    New-ConditionalText -Text 'True'        -ConditionalTextColor DarkGreen -BackgroundColor LightGreen
    New-ConditionalText -Text 'FALSE'       -ConditionalTextColor DarkRed   -BackgroundColor LightPink
    New-ConditionalText -Text 'False'       -ConditionalTextColor DarkRed   -BackgroundColor LightPink
)

# ---- 5. Process Labels and Export to Excel ---------------------------------

for ($i = 0; $i -lt $Labels.Count; $i++) {
    $label = $Labels[$i]
    # Excel sheet names: max 31 chars, no : \ / ? * [ ]
    $sheetName = ($label -replace '[:\\/?*\[\]]', '_')
    if ($sheetName.Length -gt 31) { $sheetName = $sheetName.Substring(0, 31) }
    $style = $styles[$i % $styles.Count]

    Write-Host "Fetching issues for label '$label' -> worksheet '$sheetName'..." -ForegroundColor Cyan

    $jqlQuery = "project = $ProjectKey AND labels = $label"
    $fieldsToFetch = @(
        'summary',
        'status',
        'assignee',
        'issuetype',
        'duedate',
        $StartDateFieldId,
        'description',
        $AcceptanceCriteriaFieldId
    )

    $issues = Get-JiraIssuesByJql -Jql $jqlQuery -Fields $fieldsToFetch

    $rows = @($issues | ForEach-Object {
            [pscustomobject]@{
                Key                = $_.key
                Summary            = $_.fields.summary
                Status             = $_.fields.status.name
                Assignee           = $_.fields.assignee.displayName
                IssueType          = $_.fields.issuetype.name
                StartDate          = $_.fields.$StartDateFieldId
                DueDate            = $_.fields.duedate
                Description        = (Test-HasContent $_.fields.description)
                AcceptanceCriteria = (Test-HasContent $_.fields.$AcceptanceCriteriaFieldId)
            }
        })

    Write-Host "  Found $($rows.Count) issue(s)." -ForegroundColor Gray

    if ($rows.Count -eq 0) {
        Write-Host "  No issues found for label '$label'; writing a header-only sheet." -ForegroundColor DarkYellow
        $rows = [pscustomobject]@{
            Key                = $null
            Summary            = $null
            Status             = $null
            Assignee           = $null
            IssueType          = $null
            StartDate          = $null
            DueDate            = $null
            Description        = $null
            AcceptanceCriteria = $null
        }
    }

    $rows | Export-Excel -Path $excelPath -WorksheetName $sheetName -ClearSheet `
        -AutoFilter -FreezeTopRow `
        -TableStyle $style -ConditionalText $ConditionalText
}

Write-Host "`nDone. Workbook saved to: $excelPath" -ForegroundColor Green
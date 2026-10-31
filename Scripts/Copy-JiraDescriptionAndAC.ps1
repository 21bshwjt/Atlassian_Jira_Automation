<#
.SYNOPSIS
    Copies both the Description and Acceptance Criteria (customfield_10320)
    from IDAM-21 to one or multiple Jira tickets, preserving 100% of the rich-text ADF formatting.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$JiraBaseUrl = "https://your-site.atlassian.net",

    [Parameter(Mandatory = $true, HelpMessage = "Your Atlassian account email")]
    [string]$Email,

    [Parameter(Mandatory = $true, HelpMessage = "Your Atlassian API Token")]
    [string]$ApiToken,

    [Parameter(Mandatory = $false)]
    [string]$SourceIssueKey = "IDAM-21",

    [Parameter(Mandatory = $true, ValueFromPipeline = $true, HelpMessage = "One or more target issue keys")]
    [string[]]$TargetIssueKeys,

    # Discovered from your Jira instance for "Acceptance criteria"
    [Parameter(Mandatory = $false)]
    [string]$AcceptanceCriteriaFieldId = "customfield_10320"
)

# 1. Prepare Basic Auth Headers
$authBytes = [System.Text.Encoding]::ASCII.GetBytes("${Email}:${ApiToken}")
$base64Auth = [System.Convert]::ToBase64String($authBytes)

$headers = @{
    "Authorization" = "Basic $base64Auth"
    "Content-Type"  = "application/json"
    "Accept"        = "application/json"
}

# 2. Fetch Description and Acceptance Criteria from Source Issue
try {
    $SourceIssueKey = $SourceIssueKey.Trim()
    Write-Host "Fetching details from source issue: $SourceIssueKey..." -ForegroundColor Cyan

    $fieldsParam = "description,$AcceptanceCriteriaFieldId"
    $getUrl = "{0}/rest/api/3/issue/{1}?fields={2}" -f $JiraBaseUrl.TrimEnd('/'), $SourceIssueKey, $fieldsParam

    $sourceIssue = Invoke-RestMethod -Uri $getUrl -Method Get -Headers $headers
   
    $sourceDescription = $sourceIssue.fields.description
    $sourceAcceptanceCriteria = $sourceIssue.fields.$AcceptanceCriteriaFieldId

    if ($null -eq $sourceDescription -and $null -eq $sourceAcceptanceCriteria) {
        Write-Warning "Source issue $SourceIssueKey has neither Description nor Acceptance Criteria populated."
        return
    }

    Write-Host " Retrieved content from $SourceIssueKey" -ForegroundColor Green
}
catch {
    Write-Error "Failed to retrieve source issue $SourceIssueKey : $($_.Exception.Message)"
    if ($_.Exception.Response) {
        $stream = $_.Exception.Response.GetResponseStream()
        $reader = [System.IO.StreamReader]::new($stream)
        Write-Error "API Response: $($reader.ReadToEnd())"
    }
    return
}

# 3. Construct Update Fields Payload
$fieldsToUpdate = @{}

if ($null -ne $sourceDescription) {
    $fieldsToUpdate["description"] = $sourceDescription
}

if ($null -ne $sourceAcceptanceCriteria) {
    $fieldsToUpdate[$AcceptanceCriteriaFieldId] = $sourceAcceptanceCriteria
}

$updatePayload = @{
    fields = $fieldsToUpdate
}

# CRITICAL: -Depth 100 prevents truncation of nested ADF nodes (tables, tasks, lists)
$jsonBody = $updatePayload | ConvertTo-Json -Depth 100

# 4. Resolve Target Issue Keys (handles string arrays and comma-separated inputs)
$resolvedTargets = $TargetIssueKeys | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ }

Write-Host "Ready to update $($resolvedTargets.Count) target ticket(s): $($resolvedTargets -join ', ')`n" -ForegroundColor Cyan

# 5. Update Each Target Issue
foreach ($targetKey in $resolvedTargets) {
    try {
        Write-Host "Updating $targetKey..." -NoNewline
        $putUrl = "{0}/rest/api/3/issue/{1}" -f $JiraBaseUrl.TrimEnd('/'), $targetKey

        $null = Invoke-RestMethod -Uri $putUrl -Method Put -Headers $headers -Body $jsonBody
        Write-Host " [SUCCESS]" -ForegroundColor Green
    }
    catch {
        Write-Host " [FAILED]" -ForegroundColor Red
        Write-Error "Could not update $targetKey : $($_.Exception.Message)"
        if ($_.Exception.Response) {
            $stream = $_.Exception.Response.GetResponseStream()
            $reader = [System.IO.StreamReader]::new($stream)
            Write-Error "Response body: $($reader.ReadToEnd())"
        }
    }
}

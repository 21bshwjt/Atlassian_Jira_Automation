#https://id.atlassian.com/manage-profile/security/api-tokens


function Test-JiraApiAccess {
    <#
    .SYNOPSIS
        Tests Jira Cloud REST API access using Basic auth (email + API token).

    .DESCRIPTION
        Calls /rest/api/3/myself and reports whether API access works.
        On failure, shows the HTTP status, IP-restriction / auth-denied headers,
        and the response body to help diagnose the problem.

    .PARAMETER Email
        Atlassian account email address.

    .PARAMETER ApiToken
        Atlassian API token. Use single quotes when passing it inline.

    .PARAMETER BaseUrl
        Jira Cloud base URL. Defaults to https://opella.atlassian.net

    .EXAMPLE
        Test-JiraApiAccess -Email 'username@opella.com' -ApiToken 'your-token'

    .EXAMPLE
        Test-JiraApiAccess -Email 'username@opella.com' -ApiToken 'your-token' -Verbose
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipelineByPropertyName)]
        [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]
        [string]$Email,

        [Parameter(Mandatory, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string]$ApiToken,

        [Parameter()]
        [ValidatePattern('^https://')]
        #[string]$BaseUrl = 'https://opella.atlassian.net'
        [string]$BaseUrl = 'https://orbitra.atlassian.net'
    )

    begin {
        # Helper: read a response header on both Windows PowerShell 5.1 and PowerShell 7+
        function Get-ResponseHeader {
            param($Response, [string]$Name)
            if (-not $Response -or -not $Response.Headers) { return $null }
            try {
                if ($Response.Headers -is [System.Net.WebHeaderCollection]) {
                    return $Response.Headers[$Name]
                }
                $values = $null
                if ($Response.Headers.TryGetValues($Name, [ref]$values)) {
                    return ($values -join ', ')
                }
            }
            catch { }
            return $null
        }
    }

    process {
        $uri = "$($BaseUrl.TrimEnd('/'))/rest/api/3/myself"
        $base64 = [Convert]::ToBase64String(
            [Text.Encoding]::UTF8.GetBytes("${Email}:${ApiToken}")
        )
        $headers = @{
            Authorization = "Basic $base64"
            Accept        = 'application/json'
        }

        Write-Verbose "GET $uri as $Email"

        try {
            $res = Invoke-WebRequest -Uri $uri -Headers $headers -Method Get -ErrorAction Stop

            Write-Host "API access is WORKING! HTTP $($res.StatusCode)" -ForegroundColor Green

            [pscustomobject]@{
                Success    = $true
                StatusCode = [int]$res.StatusCode
                Uri        = $uri
                Email      = $Email
            }
        }
        catch {
            $response = $_.Exception.Response
            $statusCode = if ($response) { [int]$response.StatusCode } else { $null }

            if ($null -eq $statusCode) {
                # No HTTP response at all (DNS, proxy, TLS, timeout, ...)
                Write-Host "Request failed before receiving a response: $($_.Exception.Message)" -ForegroundColor Red
            }
            else {
                Write-Host "HTTP Status Code : $statusCode" -ForegroundColor Red
            }

            # Headers that reveal IP restrictions / auth denial
            $deniedReason = Get-ResponseHeader -Response $response -Name 'X-Authentication-Denied-Reason'
            $ipBlocked = Get-ResponseHeader -Response $response -Name 'X-Aip-Blocked'
            Write-Host "Denied Reason    : $deniedReason" -ForegroundColor Yellow
            Write-Host "IP Block Header  : $ipBlocked"    -ForegroundColor Yellow

            # Response body (PS 7 exposes it via ErrorDetails; PS 5.1 via the response stream)
            $body = $_.ErrorDetails.Message
            if (-not $body -and $response) {
                try {
                    $stream = if ($response.PSObject.Methods.Name -contains 'GetResponseStream') {
                        $response.GetResponseStream()
                    }
                    else {
                        $response.Content.ReadAsStream()
                    }
                    $reader = [System.IO.StreamReader]::new($stream)
                    $body = $reader.ReadToEnd()
                    $reader.Dispose()
                }
                catch { $body = '<unable to read response body>' }
            }
            Write-Host "Response Body    : $body" -ForegroundColor Cyan

            [pscustomobject]@{
                Success      = $false
                StatusCode   = $statusCode
                DeniedReason = $deniedReason
                IpBlocked    = $ipBlocked
                Body         = $body
                Uri          = $uri
                Email        = $Email
            }
        }
    }
}

# Example:
# Test-JiraApiAccess -Email 'username@opella.com' -ApiToken 'your-api-token'

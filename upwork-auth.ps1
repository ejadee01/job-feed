# One-time Upwork API setup for the job feed. Run it yourself in PowerShell:
#   powershell -ExecutionPolicy Bypass -File upwork-auth.ps1
# It signs you in to Upwork in your browser, swaps the sign-in code for tokens, and saves
# UPWORK_CLIENT_ID, UPWORK_CLIENT_SECRET and UPWORK_REFRESH_TOKEN as GitHub secrets of this repo.
# Nothing is written to disk. Needs the GitHub CLI (gh) signed in to an account that owns the repo.
# Run it again any time the daily workflow warns that the Upwork token refresh failed.

$Repo = 'ejadee01/job-feed'
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

function Read-Secret($prompt) {
    $s = Read-Host $prompt -AsSecureString
    [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($s))
}

$clientId     = (Read-Host 'Upwork API Key (client_id)').Trim()
$clientSecret = (Read-Secret 'Upwork API Secret (client_secret)').Trim()
$redirect     = (Read-Host 'Callback URL exactly as entered on the API key [https://ejadee01.github.io/job-feed/]').Trim()
if (-not $redirect) { $redirect = 'https://ejadee01.github.io/job-feed/' }

$auth = 'https://www.upwork.com/ab/account-security/oauth2/authorize?response_type=code&client_id=' +
        [uri]::EscapeDataString($clientId) + '&redirect_uri=' + [uri]::EscapeDataString($redirect)
Write-Host "`nOpening Upwork in your browser. Sign in and click Allow."
Write-Host "Upwork then sends you to your callback URL. Copy the whole address from the address bar (it contains ?code=...)."
Start-Process $auth

$back = (Read-Host "`nPaste the address (or just the code)").Trim()
$code = if ($back -match '[?&]code=([^&#]+)') { [uri]::UnescapeDataString($Matches[1]) } else { $back }
if (-not $code) { throw 'No code found.' }

$tok = Invoke-RestMethod -Method Post -Uri 'https://www.upwork.com/api/v3/oauth2/token' -Headers @{ Accept = 'application/json' } `
    -ContentType 'application/x-www-form-urlencoded' -Body @{ grant_type = 'authorization_code'; client_id = $clientId
        client_secret = $clientSecret; code = $code; redirect_uri = $redirect }
if (-not $tok.refresh_token) { throw 'Upwork did not return a refresh token.' }

# Quick test: one small job search, to confirm the key has the "Read marketplace Job Postings" permission
$test = @{ query = 'query($f: MarketplaceJobPostingsSearchFilter) { marketplaceJobPostingsSearch(marketPlaceJobFilter: $f, searchType: USER_JOBS_SEARCH) { totalCount } }'
           variables = @{ f = @{ titleExpression_eq = 'controller'; pagination_eq = @{ after = '0'; first = 1 } } } } | ConvertTo-Json -Depth 6 -Compress
$r = Invoke-RestMethod -Method Post -Uri 'https://api.upwork.com/graphql' -ContentType 'application/json' `
    -Headers @{ Authorization = "Bearer $($tok.access_token)" } -Body $test
if ($r.errors) { Write-Warning "Signed in, but the test search failed: $($r.errors.message -join '; '). Check that the API key has the 'Read marketplace Job Postings' permission." }
else { Write-Host "Test search OK: $($r.data.marketplaceJobPostingsSearch.totalCount) controller jobs on Upwork right now." }

gh secret set UPWORK_CLIENT_ID     --repo $Repo --body $clientId
gh secret set UPWORK_CLIENT_SECRET --repo $Repo --body $clientSecret
gh secret set UPWORK_REFRESH_TOKEN --repo $Repo --body $tok.refresh_token

$pat = (Read-Secret "`nOptional: paste a GitHub fine-grained token with 'Secrets: Read and write' on $Repo (Enter to skip)").Trim()
if ($pat) { gh secret set SECRETS_PAT --repo $Repo --body $pat }

Write-Host "`nDone. Upwork jobs will be included from the next daily run (or Actions > Refresh job feed > Run workflow)."

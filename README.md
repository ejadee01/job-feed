# Job feed

Daily list of CFO / Controller / Accountant jobs that are remote (open to the Philippines) or in Davao City, paying at least PHP 80,000/month or equivalent.

- `jobs.html` — browsable page with search and filters, live at https://ejadee01.github.io/job-feed/
- `jobs-feed.xml` — the same jobs as an RSS feed: https://ejadee01.github.io/job-feed/jobs-feed.xml
- `build-job-feed.ps1` — builds both files; settings (minimum pay, keywords, exchange rates) are at the top

- `upwork-auth.ps1` — one-time Upwork API sign-in (see below)

Sources: JobStreet, LinkedIn, Kalibrr, OnlineJobs.ph, Upwork (official API, once set up), Himalayas, Jobicy, Remotive, Remote OK. Indeed blocks automated access and has no public job-search API, so it is not included.

A GitHub Action (`.github/workflows/refresh.yml`) rebuilds the files every day at 7:00 AM Philippine time. To run it now: Actions tab → Refresh job feed → Run workflow.

## Upwork setup

Upwork jobs come from Upwork's official GraphQL API. Until the steps below are done, the build skips Upwork and everything else works as before.

1. **Request an API key** at https://www.upwork.com/developer/keys/apply (signed in to your Upwork freelancer account). Choose key type **OAuth 2.0**, set the callback URL to `https://ejadee01.github.io/job-feed/`, and tick the **Read marketplace Job Postings** permission. Upwork reviews each request by hand, so approval can take a few days.
2. **Optional but recommended: create a GitHub token** so the daily run can save the new Upwork token it receives. GitHub → Settings → Developer settings → Fine-grained tokens → Generate. Repository access: only `job-feed`. Permissions: **Secrets: Read and write**. Copy the token.
3. **Run the setup script** in PowerShell from this folder and follow the prompts:
   ```
   powershell -ExecutionPolicy Bypass -File upwork-auth.ps1
   ```
   It opens Upwork in your browser, asks you to paste back the address Upwork sends you to, runs a test search, and stores everything as GitHub secrets (`UPWORK_CLIENT_ID`, `UPWORK_CLIENT_SECRET`, `UPWORK_REFRESH_TOKEN`, and `SECRETS_PAT` if you pasted the token from step 2).

Upwork's sign-in lasts 24 hours and its refresh token expires after 2 weeks without use. The daily run renews it each morning. If a run's log says the Upwork token refresh failed (for example after the workflow was paused for more than 2 weeks), run `upwork-auth.ps1` again.

Upwork pay: hourly jobs show their maximum hourly rate converted to a monthly figure (160 hours). Fixed-price budgets are one-off amounts, so they appear in the description and the job counts as "salary not stated". Jobs whose client only accepts freelancers from other countries are skipped.

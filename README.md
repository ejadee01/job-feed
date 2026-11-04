# Job feed

Daily list of CFO / Controller / Accountant jobs that are remote (open to the Philippines) or in Davao City, paying at least PHP 80,000/month or equivalent.

- `jobs.html` — browsable page with search and filters (download it and open in a browser)
- `jobs-feed.xml` — the same jobs as an RSS feed
- `build-job-feed.ps1` — builds both files; settings (minimum pay, keywords, exchange rates) are at the top

Sources: JobStreet, LinkedIn, Kalibrr, Himalayas, Jobicy, Remotive, Remote OK. Indeed blocks automated access, so it is not included.

A GitHub Action (`.github/workflows/refresh.yml`) rebuilds the files every day at 7:00 AM Philippine time. To run it now: Actions tab → Refresh job feed → Run workflow.

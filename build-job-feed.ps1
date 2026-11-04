# Builds an RSS feed of CFO / Controller / Accountant jobs that are remote (open to the
# Philippines) or based in Davao City, paying at least the minimum monthly PHP rate.
# Run:  powershell -ExecutionPolicy Bypass -File build-job-feed.ps1
# Output: jobs-feed.xml next to this script.

# ---------------- Settings ----------------
$MinMonthlyPhp   = 80000
$IncludeNoSalary = $true    # most PH listings hide the salary; set $false to keep only confirmed matches
$TitlePattern    = '\b(CFO|chief financial officer|financial controller|controller|comptroller|accountant|accounting manager|head of accounting|finance director)\b'
$ExcludePattern  = 'air traffic|traffic controller|quality controller|document controller|inventory controller|project controller|game|controller \(m/w/d\)'
$LocationPattern = 'philippines|davao|worldwide|anywhere|global|apac|asia|southeast asia'
$OutFile         = Join-Path $PSScriptRoot 'jobs-feed.xml'

# Approximate PHP per 1 unit of currency - update occasionally
$FxToPhp = @{ PHP = 1; USD = 58; AUD = 38; CAD = 42; EUR = 63; GBP = 75; SGD = 44; NZD = 34; HKD = 7.4 }

# ---------------- Helpers ----------------
$wc = New-Object Net.WebClient
$wc.Encoding = [Text.Encoding]::UTF8
function Get-Json($url) {
    try {
        $wc.Headers['User-Agent'] = 'Mozilla/5.0 (personal job feed)'
        return $wc.DownloadString($url) | ConvertFrom-Json
    } catch { Write-Warning "Failed: $url - $($_.Exception.Message)"; return $null }
}

# Returns the monthly PHP equivalent of a salary figure, or $null if unknown.
function ConvertTo-MonthlyPhp($amount, $currency, $period) {
    if (-not $amount -or [double]$amount -le 0) { return $null }
    $cur = if ($currency) { "$currency".ToUpper() } else { 'USD' }
    if (-not $FxToPhp.ContainsKey($cur)) { return $null }
    $php = [double]$amount * $FxToPhp[$cur]
    switch -Regex ("$period") {
        'hour'  { return $php * 160 }
        'week'  { return $php * 4.33 }
        'month' { return $php }
        default {
            # Annual (or unlabeled). Tiny annual figures are really monthly figures mislabeled.
            if ($php / 12 -lt 15000) { return $php } else { return $php / 12 }
        }
    }
}

$jobs = New-Object System.Collections.Generic.List[object]
function Add-Job($title, $company, $link, $location, $date, $source, $salMin, $salMax, $currency, $period, $summary) {
    $monthly = ConvertTo-MonthlyPhp ($(if ($salMax) { $salMax } else { $salMin })) $currency $period
    $jobs.Add([pscustomobject]@{
        Title = "$title".Trim(); Company = "$company".Trim(); Link = $link; Location = "$location"
        Date = $date; Source = $source; MonthlyPhp = $monthly; Summary = $summary
        SalaryText = if ($salMin -or $salMax) { "$salMin-$salMax $currency $period" } else { '' }
    })
}

# ---------------- Sources ----------------
# Himalayas: remote jobs restricted to / open to the Philippines
foreach ($q in 'accountant', 'controller', 'CFO', 'chief financial officer', 'accounting manager') {
    $r = Get-Json ("https://himalayas.app/jobs/api/search?country=PH&limit=50&q=" + [uri]::EscapeDataString($q))
    foreach ($j in $r.jobs) {
        $loc = if ($j.locationRestrictions) { $j.locationRestrictions -join ', ' } else { 'Worldwide' }
        Add-Job $j.title $j.companyName $j.applicationLink "Remote - $loc" ([DateTimeOffset]::FromUnixTimeSeconds([int64]$j.pubDate).UtcDateTime) 'Himalayas' $j.minSalary $j.maxSalary $j.currency $j.salaryPeriod $j.excerpt
    }
}

# Jobicy: accounting & finance, Philippines-eligible and APAC
foreach ($geo in 'philippines', 'apac') {
    $r = Get-Json "https://jobicy.com/api/v2/remote-jobs?count=100&industry=accounting-finance&geo=$geo"
    foreach ($j in $r.jobs) {
        Add-Job $j.jobTitle $j.companyName $j.url "Remote - $($j.jobGeo)" ([datetime]$j.pubDate) 'Jobicy' $j.salaryMin $j.salaryMax $j.salaryCurrency $j.salaryPeriod $j.jobExcerpt
    }
}

# Remotive: finance & legal category (location filtered below)
$r = Get-Json 'https://remotive.com/api/remote-jobs?category=finance-legal'
foreach ($j in $r.jobs) {
    $min = $null; $max = $null; $cur = 'USD'; $per = 'year'
    if ($j.salary -match '([\d.,]+)\s*(k)?') {
        $nums = [regex]::Matches($j.salary, '([\d.,]+)\s*(k)?') | % { [double]($_.Groups[1].Value -replace ',', '') * $(if ($_.Groups[2].Value) { 1000 } else { 1 }) }
        $min = $nums | select -First 1; $max = $nums | select -Last 1
        if ($j.salary -match 'month') { $per = 'month' } elseif ($j.salary -match 'hour') { $per = 'hour' }
        if ($j.salary -match '€') { $cur = 'EUR' } elseif ($j.salary -match '£') { $cur = 'GBP' }
    }
    Add-Job $j.title $j.company_name $j.url "Remote - $($j.candidate_required_location)" ([datetime]$j.publication_date) 'Remotive' $min $max $cur $per ''
}

# Remote OK: all jobs (salary in USD/year, location filtered below)
$r = Get-Json 'https://remoteok.com/api'
foreach ($j in ($r | Select-Object -Skip 1)) {
    $loc = if ($j.location) { $j.location } else { 'Worldwide' }
    Add-Job $j.position $j.company $j.url "Remote - $loc" ([datetime]$j.date) 'Remote OK' $j.salary_min $j.salary_max 'USD' 'year' ''
}

# Parses a salary label like "P80,000 - P100,000 per month" into min, max, currency, period.
function Read-SalaryLabel($text) {
    if (-not $text) { return @($null, $null, $null, $null) }
    $nums = [regex]::Matches($text, '(\d[\d,]*(?:\.\d+)?)\s*(k)?', 'IgnoreCase') | % {
        [double]($_.Groups[1].Value -replace ',', '') * $(if ($_.Groups[2].Value) { 1000 } else { 1 }) }
    if (-not $nums) { return @($null, $null, $null, $null) }
    $cur = if ($text -match '\$|USD') { 'USD' } else { 'PHP' }
    $per = if ($text -match 'hour') { 'hour' } elseif ($text -match 'year|annum|annual') { 'year' } else { 'month' }
    return @(($nums | select -First 1), ($nums | select -Last 1), $cur, $per)
}

$Keywords = 'accountant', 'controller', 'CFO', 'chief financial officer', 'accounting manager'

# JobStreet PH (site's own search API): Davao City jobs plus nationwide remote jobs
foreach ($q in $Keywords) {
    foreach ($where in 'Davao City', '') {
        $url = "https://ph.jobstreet.com/api/jobsearch/v5/search?siteKey=PH-Main&pageSize=100&page=1&keywords=" + [uri]::EscapeDataString($q)
        if ($where) { $url += '&where=' + [uri]::EscapeDataString($where) }
        $r = Get-Json $url
        foreach ($j in $r.data) {
            $place = $j.locations.label -join '; '
            $setup = $j.workArrangements.data.label.text -join '; '
            $loc = if ($setup -match 'Remote') { "Remote - Philippines ($place)" } else { "$place ($setup)" }
            $s = Read-SalaryLabel $j.salaryLabel
            Add-Job $j.title $(if ($j.companyName) { $j.companyName } else { $j.advertiser.description }) "https://ph.jobstreet.com/job/$($j.id)" $loc ([datetime]$j.listingDate) 'JobStreet' $s[0] $s[1] $s[2] $s[3] $j.teaser
        }
    }
}

# Kalibrr (site's own search API): keep Davao City and work-from-home jobs
foreach ($q in $Keywords) {
    $r = Get-Json ("https://www.kalibrr.com/kjs/job_board/search?limit=100&offset=0&text=" + [uri]::EscapeDataString($q))
    foreach ($j in $r.jobs) {
        $city = $j.google_location.address_components.city
        $loc = if ($j.is_work_from_home) { "Remote - Philippines ($city)" } else { "$city" }
        $per = if ("$($j.salary_interval)" -match 'year|annual') { 'year' } elseif ("$($j.salary_interval)" -match 'hour') { 'hour' } else { 'month' }
        $cur = if ($j.salary_currency) { $j.salary_currency } else { 'PHP' }
        Add-Job $j.name $j.company_name "https://www.kalibrr.com/c/$($j.company.code)/jobs/$($j.id)/$($j.slug)" $loc ([datetime]$j.activation_date) 'Kalibrr' $j.base_salary $j.maximum_salary $cur $per ''
    }
}

# LinkedIn (public guest job search, no login): Davao City jobs plus remote jobs in the Philippines
foreach ($q in $Keywords) {
    foreach ($p in '&location=Davao%20City%2C%20Philippines', '&location=Philippines&f_WT=2') {
        $url = "https://www.linkedin.com/jobs-guest/jobs/api/seeMoreJobPostings/search?keywords=" + [uri]::EscapeDataString($q) + $p
        Start-Sleep -Seconds 3
        try { $wc.Headers['User-Agent'] = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/130.0'; try { $html = $wc.DownloadString($url) } catch { Start-Sleep -Seconds 20; $html = $wc.DownloadString($url) } }
        catch { Write-Warning "Failed: $url - $($_.Exception.Message)"; continue }
        foreach ($card in ($html -split '<li>' | Select-Object -Skip 1)) {
            $get = { param($pat) $m = [regex]::Match($card, $pat, 'Singleline'); [Net.WebUtility]::HtmlDecode($m.Groups[1].Value.Trim()) }
            $title = & $get 'base-search-card__title">(.*?)</h3>'
            if (-not $title) { continue }
            $link = (& $get 'base-card__full-link[^>]*href="([^"?]+)')
            $place = & $get 'job-search-card__location">(.*?)</span>'
            $loc = if ($p -match 'f_WT=2') { "Remote - Philippines ($place)" } else { $place }
            $date = & $get 'datetime="([^"]+)"'
            $s = Read-SalaryLabel (& $get 'job-search-card__salary-info">(.*?)</span>')
            Add-Job $title ((& $get 'base-search-card__subtitle">(.*?)</h4>') -replace '<[^>]+>', '' -replace '\s+', ' ').Trim() $link $loc $(if ($date) { [datetime]$date } else { Get-Date }) 'LinkedIn' $s[0] $s[1] $s[2] $s[3] ''
        }
    }
}

# OnlineJobs.ph: remote jobs for Filipino workers (robots.txt asks for a 5-second crawl delay)
foreach ($q in $Keywords) {
    foreach ($offset in '', '/30') {
        Start-Sleep -Seconds 5
        $url = "https://www.onlinejobs.ph/jobseekers/jobsearch$offset" + '?jobkeyword=' + [uri]::EscapeDataString($q)
        try { $wc.Headers['User-Agent'] = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/130.0'; $html = $wc.DownloadString($url) }
        catch { Write-Warning "Failed: $url - $($_.Exception.Message)"; continue }
        foreach ($card in ($html -split '<a href="/jobseekers/job/' | Select-Object -Skip 1)) {
            if ($card -notmatch 'jobpost-cat-box') { continue }
            $slug = ($card -split '"')[0]
            $m = [regex]::Match($card, '<h4[^>]*>(.*?)<span[^>]*>(.*?)</span>', 'Singleline')
            $title = [Net.WebUtility]::HtmlDecode($m.Groups[1].Value.Trim()); $type = $m.Groups[2].Value.Trim()
            if (-not $title) { continue }
            $date = [regex]::Match($card, 'data-temp-2="([^"]+)"').Groups[1].Value
            $pay = [Net.WebUtility]::HtmlDecode([regex]::Match($card, 'icon-round-dollar.*?<dd[^>]*>(.*?)</dd>', 'Singleline').Groups[1].Value.Trim())
            $desc = ([regex]::Match($card, 'class="desc[^"]*">(.*?)<a ', 'Singleline').Groups[1].Value -replace '<[^>]+>', ' ' -replace '\s+', ' ').Trim()
            # Salary is free text: no currency usually means USD/month; big numbers are pesos; tiny ones are hourly
            $s = Read-SalaryLabel $pay
            if ($s[0] -and $pay -notmatch '\$|USD|₱|PHP|peso|php') {
                $s[2] = if ($s[1] -ge 10000) { 'PHP' } else { 'USD' }
                if ($s[1] -lt 30 -or $pay -match 'hr|hour') { $s[3] = 'hour' }
            } elseif ($s[0] -and $pay -match '₱|PHP|peso') { $s[2] = 'PHP' }
            $when = if ($date) { [datetime]::SpecifyKind([datetime]$date, 'Utc') } else { Get-Date }
            Add-Job $title "OnlineJobs.ph employer ($type)" "https://www.onlinejobs.ph/jobseekers/job/$slug" 'Remote - Philippines' $when 'OnlineJobs.ph' $s[0] $s[1] $s[2] $s[3] $desc
        }
    }
}

# ---------------- Filter ----------------
$seen = @{}
$selected = $jobs | Where-Object {
    $_.Title -match $TitlePattern -and $_.Title -notmatch $ExcludePattern -and
    $_.Location -match $LocationPattern -and
    ($(if ($null -eq $_.MonthlyPhp) { $IncludeNoSalary } else { $_.MonthlyPhp -ge $MinMonthlyPhp }))
} | Where-Object {
    $key = ($_.Title + '|' + $_.Company).ToLower()
    if ($seen[$key]) { $false } else { $seen[$key] = $true; $true }
} | Sort-Object Date -Descending

# ---------------- Write RSS ----------------
function Esc($s) { [Security.SecurityElement]::Escape("$s") }
$sb = New-Object Text.StringBuilder
[void]$sb.AppendLine('<?xml version="1.0" encoding="UTF-8"?>')
[void]$sb.AppendLine('<rss version="2.0"><channel>')
[void]$sb.AppendLine('<title>CFO / Controller / Accountant jobs - Remote (PH) or Davao</title>')
[void]$sb.AppendLine('<link>https://himalayas.app/jobs/countries/philippines</link>')
[void]$sb.AppendLine("<description>Min PHP $('{0:N0}' -f $MinMonthlyPhp)/month or equivalent. Sources: JobStreet, LinkedIn, Kalibrr, OnlineJobs.ph, Himalayas, Jobicy, Remotive, Remote OK.</description>")
[void]$sb.AppendLine("<lastBuildDate>$((Get-Date).ToUniversalTime().ToString('r'))</lastBuildDate>")
foreach ($m in $selected) {
    $pay = if ($m.MonthlyPhp) { "~PHP {0:N0}/mo" -f $m.MonthlyPhp } else { 'Salary not stated' }
    $desc = "<p><b>$pay</b> | $($m.Location) | via $($m.Source)</p>"
    if ($m.SalaryText) { $desc += "<p>Posted salary: $($m.SalaryText)</p>" }
    if ($m.Summary) { $desc += "<p>$($m.Summary)</p>" }
    [void]$sb.AppendLine('<item>')
    [void]$sb.AppendLine("<title>$(Esc "[$pay] $($m.Title) - $($m.Company)")</title>")
    [void]$sb.AppendLine("<link>$(Esc $m.Link)</link>")
    [void]$sb.AppendLine("<guid isPermaLink=`"true`">$(Esc $m.Link)</guid>")
    [void]$sb.AppendLine("<pubDate>$(([datetime]$m.Date).ToUniversalTime().ToString('r'))</pubDate>")
    [void]$sb.AppendLine("<description>$(Esc $desc)</description>")
    [void]$sb.AppendLine('</item>')
}
[void]$sb.AppendLine('</channel></rss>')
[IO.File]::WriteAllText($OutFile, $sb.ToString(), (New-Object Text.UTF8Encoding $false))

# ---------------- Write HTML page ----------------
$HtmlFile = Join-Path $PSScriptRoot 'jobs.html'
$data = @($selected | % {
    [ordered]@{ t = $_.Title; c = $_.Company; l = $_.Link; loc = $_.Location; s = $_.Source
                d = ([datetime]$_.Date).ToString('yyyy-MM-dd'); p = $(if ($_.MonthlyPhp) { [math]::Round($_.MonthlyPhp) } else { $null })
                raw = $_.SalaryText }
})
$json = (ConvertTo-Json -InputObject $data -Depth 3 -Compress) -replace '</', '<\/'
$template = @'
<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Finance Job Board</title>
<style>
:root{--bg:#f7f7f5;--card:#fff;--text:#1d1d1b;--muted:#6b6b66;--line:#e4e3de;--accent:#1f6f5c;--pay:#e6f3ee;--nopay:#f0efeb}
@media (prefers-color-scheme:dark){:root{--bg:#161615;--card:#1f1f1d;--text:#ecebe6;--muted:#9a9993;--line:#33322f;--accent:#5cc2a4;--pay:#1e3a31;--nopay:#2a2a27}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--text);font:15px/1.5 system-ui,-apple-system,"Segoe UI",sans-serif}
main{max-width:960px;margin:0 auto;padding:24px 16px 48px}
h1{font-size:22px;margin:0 0 4px}.sub{color:var(--muted);margin:0 0 20px;font-size:14px}
.bar{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:16px}
.bar input[type=search],.bar select{padding:8px 10px;border:1px solid var(--line);border-radius:8px;background:var(--card);color:var(--text);font:inherit}
.bar input[type=search]{flex:1 1 220px}.bar label{display:flex;align-items:center;gap:6px;font-size:14px;color:var(--muted)}
.count{font-size:13px;color:var(--muted);margin-bottom:8px}
.job{display:block;background:var(--card);border:1px solid var(--line);border-radius:10px;padding:14px 16px;margin-bottom:10px;color:inherit;text-decoration:none}
.job:hover{border-color:var(--accent)}
.top{display:flex;justify-content:space-between;gap:12px;align-items:flex-start}
.title{font-weight:600}.co{color:var(--muted);font-size:14px}
.pay{white-space:nowrap;font-size:13px;font-weight:600;padding:3px 8px;border-radius:6px;background:var(--pay);color:var(--accent)}
.pay.none{background:var(--nopay);color:var(--muted);font-weight:500}
.meta{margin-top:6px;font-size:13px;color:var(--muted);display:flex;flex-wrap:wrap;gap:4px 14px}
@media (max-width:560px){.top{flex-direction:column}}
</style></head><body><main>
<h1>CFO / Controller / Accountant jobs</h1>
<p class="sub">Remote (open to the Philippines) or Davao City &middot; min PHP __MIN__/month or equivalent &middot; updated __UPDATED__</p>
<div class="bar">
<input type="search" id="q" placeholder="Search title, company, location">
<select id="src"><option value="">All sources</option></select>
<select id="where"><option value="">Remote + Davao</option><option value="remote">Remote only</option><option value="davao">Davao only</option></select>
<label><input type="checkbox" id="paid"> Salary stated only</label>
</div>
<div class="count" id="count"></div><div id="list"></div>
</main><script>
const jobs = __DATA__;
const $ = id => document.getElementById(id);
[...new Set(jobs.map(j => j.s))].sort().forEach(s => $('src').add(new Option(s, s)));
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
function render() {
  const q = $('q').value.toLowerCase(), src = $('src').value, w = $('where').value, paid = $('paid').checked;
  const rows = jobs.filter(j =>
    (!q || (j.t + ' ' + j.c + ' ' + j.loc).toLowerCase().includes(q)) && (!src || j.s === src) &&
    (!paid || j.p) && (!w || (w === 'remote' ? /^Remote/.test(j.loc) : /davao/i.test(j.loc))));
  $('count').textContent = rows.length + ' of ' + jobs.length + ' jobs';
  $('list').innerHTML = rows.map(j => `<a class="job" href="${esc(j.l)}" target="_blank" rel="noopener">
    <div class="top"><div><div class="title">${esc(j.t)}</div><div class="co">${esc(j.c)}</div></div>
    <span class="pay${j.p ? '' : ' none'}" title="${esc(j.raw)}">${j.p ? '~PHP ' + j.p.toLocaleString() + '/mo' : 'Salary not stated'}</span></div>
    <div class="meta"><span>${esc(j.loc)}</span><span>${esc(j.s)}</span><span>${esc(j.d)}</span></div></a>`).join('');
}
['q','src','where','paid'].forEach(id => $(id).addEventListener('input', render));
render();
</script></body></html>
'@
$html = $template.Replace('__DATA__', $json).Replace('__MIN__', ('{0:N0}' -f $MinMonthlyPhp)).Replace('__UPDATED__', ([DateTime]::UtcNow.AddHours(8).ToString('MMM d, yyyy h:mm tt') + ' PH time'))
[IO.File]::WriteAllText($HtmlFile, $html, (New-Object Text.UTF8Encoding $false))

$withPay = @($selected | ? { $_.MonthlyPhp }).Count
Write-Host "Wrote $(@($selected).Count) jobs ($withPay with confirmed salary >= PHP $MinMonthlyPhp) to $OutFile"

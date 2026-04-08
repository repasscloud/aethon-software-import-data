[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$SearchUrl,

    [string]$BaseUrl = "https://careers-inc.nttdata.com",

    [Parameter(Mandatory)]
    [string]$OutputPath
)
# param(
#     [string]$SearchUrl   = "https://careers-inc.nttdata.com/search/?q=&locationsearch=PH&location=PH&sortColumn=referencedate&sortDirection=desc",
#     [string]$BaseUrl     = "https://careers-inc.nttdata.com",
#     [string]$OutputPath  = ".\nttdata-jobs.json"
# )

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-FirstMatch {
    param(
        [Parameter(Mandatory)]
        [string]$InputString,

        [Parameter(Mandatory)]
        [string]$Pattern
    )

    $m = [regex]::Match(
        $InputString,
        $Pattern,
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )

    if ($m.Success) {
        return $m.Groups[1].Value
    }

    return $null
}

function Invoke-StripHtml {
    param(
        [AllowNull()]
        [string]$Html
    )

    if ([string]::IsNullOrWhiteSpace($Html)) {
        return $null
    }

    $text = $Html

    # preserve some structure before stripping tags
    $text = [regex]::Replace($text, '<br\s*/?>', "`n", 'IgnoreCase')
    $text = [regex]::Replace($text, '</p\s*>', "`n", 'IgnoreCase')
    $text = [regex]::Replace($text, '</li\s*>', "`n", 'IgnoreCase')

    # remove remaining tags
    $text = [regex]::Replace($text, '<[^>]+>', ' ')

    # decode entities and normalise whitespace
    $text = [System.Net.WebUtility]::HtmlDecode($text)
    $text = $text -replace "`r", ''
    $text = [regex]::Replace($text, "[`t ]+", " ")
    $text = [regex]::Replace($text, "\n\s+", "`n")
    $text = [regex]::Replace($text, "\n{2,}", "`n")
    $text = $text.Trim()

    return $text
}

function Invoke-NormalizeHtmlFragment {
    param(
        [AllowNull()]
        [string]$Html
    )

    if ([string]::IsNullOrWhiteSpace($Html)) {
        return $null
    }

    $normalized = $Html.Trim()

    # decode HTML entities
    $normalized = [System.Net.WebUtility]::HtmlDecode($normalized)

    # normalise whitespace between tags and inside content
    $normalized = $normalized -replace "`r", ''
    $normalized = $normalized -replace "`n", ' '
    $normalized = $normalized -replace "`t", ' '
    $normalized = [regex]::Replace($normalized, '>\s+<', '><')
    $normalized = [regex]::Replace($normalized, '\s{2,}', ' ')
    $normalized = $normalized.Trim()

    return $normalized
}

function Convert-DateTextToYyyyMMdd {
    param(
        [AllowNull()]
        [string]$DateText
    )

    if ([string]::IsNullOrWhiteSpace($DateText)) {
        return $null
    }

    $clean = $DateText.Trim()

    # example: Apr 8, 2026
    $formats = @(
        "MMM d, yyyy",
        "MMM dd, yyyy",
        "MMMM d, yyyy",
        "MMMM dd, yyyy"
    )

    foreach ($format in $formats) {
        [datetime]$parsed = [datetime]::MinValue
        if ([datetime]::TryParseExact(
            $clean,
            $format,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::None,
            [ref]$parsed
        )) {
            return $parsed.ToString("yyyyMMdd")
        }
    }

    return $null
}

function Get-AbsoluteUrl {
    param(
        [AllowNull()]
        [string]$BaseUrl,

        [AllowNull()]
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $null
    }

    if ($Path -match '^https?://') {
        return $Path
    }

    return ([System.Uri]::new([System.Uri]$BaseUrl, $Path)).AbsoluteUri
}

# 1. Load search page
$searchResponse = Invoke-WebRequest -Uri $SearchUrl

# 2. Extract unique relative job URLs
$jobPaths = [regex]::Matches($searchResponse.Content, '/job/[^/]+/\d+/') | ForEach-Object {
    $_.Value
} | Sort-Object -Unique

# 3. Visit each job URL and extract fields
$results = New-Object System.Collections.Generic.List[object]

foreach ($jobPath in $jobPaths) {
    $jobUrl = Get-AbsoluteUrl -BaseUrl $BaseUrl -Path $jobPath

    try {
        $jobResponse = Invoke-WebRequest -Uri $jobUrl
        $html = $jobResponse.Content

        # og:title
        $ogTitle = Get-FirstMatch -InputString $html -Pattern '<meta\s+property="og:title"\s+content="(.*?)"\s*/?>'
        if ($null -ne $ogTitle) {
            $ogTitle = [System.Net.WebUtility]::HtmlDecode($ogTitle).Trim()
        }

        # Date from <p id="job-date">
        $jobDateInner = Get-FirstMatch -InputString $html -Pattern '<p[^>]*id="job-date"[^>]*>(.*?)</p>'
        $jobDateText = Invoke-StripHtml $jobDateInner
        if ($null -ne $jobDateText) {
            $jobDateText = ($jobDateText -replace '^\s*Date:\s*', '').Trim()
        }
        $jobDateYyyyMMdd = Convert-DateTextToYyyyMMdd $jobDateText

        # Country code from jobGeoLocation => last 2 letters
        $geoInner = Get-FirstMatch -InputString $html -Pattern '<span\s+class="jobGeoLocation">(.*?)</span>'
        $geoText = Invoke-StripHtml $geoInner
        $countryCode = $null
        if (-not [string]::IsNullOrWhiteSpace($geoText)) {
            $mCountry = [regex]::Match($geoText.Trim(), '([A-Z]{2})\s*$')
            if ($mCountry.Success) {
                $countryCode = $mCountry.Groups[1].Value
            }
        }

        # Description HTML inside <span itemprop="description" class="jobdescription">...</span>
        $descriptionInnerHtml = Get-FirstMatch -InputString $html -Pattern '<span\s+itemprop="description"\s+class="jobdescription">(.*?)</span>\s*<p\s+class="job-location">'
        $descriptionHtmlNormalized = Invoke-NormalizeHtmlFragment $descriptionInnerHtml
        $descriptionText = Invoke-StripHtml $descriptionInnerHtml
        if ($null -ne $descriptionText) {
            $descriptionText = [regex]::Replace($descriptionText, "\n", " ")
            $descriptionText = [regex]::Replace($descriptionText, "\s{2,}", " ").Trim()
        }

        # Job segment / industry
        $industryInner = Get-FirstMatch -InputString $html -Pattern '<span\s+itemprop="industry">(.*?)</span>'
        $industryText = Invoke-StripHtml $industryInner
        $jobSegment = $null
        if (-not [string]::IsNullOrWhiteSpace($industryText)) {
            $jobSegment = [regex]::Replace($industryText.Trim(), '^Job\s+Segment:\s*', '', 'IgnoreCase').Trim()
        }

        # Apply URL
        $applyPath = Get-FirstMatch -InputString $html -Pattern '<a[^>]*class="[^"]*\bapply\s+dialogApplyBtn\b[^"]*"[^>]*href="([^"]+)"[^>]*>'
        $applyUrl = Get-AbsoluteUrl -BaseUrl $BaseUrl -Path $applyPath

        # Optional job ID from URL
        $jobId = $null
        $mJobId = [regex]::Match($jobPath, '/job/[^/]+/(\d+)/')
        if ($mJobId.Success) {
            $jobId = $mJobId.Groups[1].Value
        }

        $results.Add([PSCustomObject]@{
            jobId                   = $jobId
            jobUrl                  = $jobUrl
            title                   = $ogTitle
            datePostedRaw           = $jobDateText
            datePosted              = $jobDateYyyyMMdd
            countryCode             = $countryCode
            descriptionHtml         = $descriptionHtmlNormalized
            descriptionText         = $descriptionText
            jobSegment              = $jobSegment
            applyUrl                = $applyUrl
        })
    }
    catch {
        Write-Warning "Skipping $jobUrl because it failed: $($_.Exception.Message)"
        continue
    }
}

# 4. Write JSON
$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
$json = $results | ConvertTo-Json -Depth 10
[System.IO.File]::WriteAllText($OutputPath, $json, $utf8NoBom)

# 5. Output to pipeline as well
$results

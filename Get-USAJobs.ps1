[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path -Path $PSScriptRoot -ChildPath "usajobs-import.json")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$USAJOBS_API_KEY    = $env:USAJOBS_API_KEY
$USAJOBS_USER_AGENT = $env:USAJOBS_USER_AGENT
$COMPANY_LOGO_URL   = "https://raw.githubusercontent.com/repasscloud/aethon-software-import-data/refs/heads/main/usajobs/img/icons/red-2x.svg"

# ---------------------------------------------------------------------------
# API fetch
# ---------------------------------------------------------------------------

function Get-USAJobsSearchResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ApiKey,

        [Parameter(Mandatory)]
        [string]$UserAgentEmail,

        [int]$Page = 1,

        [ValidateRange(1, 500)]
        [int]$ResultsPerPage = 500
    )

    $uri = "https://data.usajobs.gov/api/Search?Page=$Page&ResultsPerPage=$ResultsPerPage"
    $client = [System.Net.Http.HttpClient]::new()

    try {
        $client.DefaultRequestHeaders.TryAddWithoutValidation("Host", "data.usajobs.gov") | Out-Null
        $client.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", $UserAgentEmail) | Out-Null
        $client.DefaultRequestHeaders.TryAddWithoutValidation("Authorization-Key", $ApiKey) | Out-Null

        $json = $client.GetStringAsync($uri).GetAwaiter().GetResult()
        $response = $json | ConvertFrom-Json

        $searchResult = $response.SearchResult

        return [PSCustomObject]@{
            Page            = $Page
            ResultsPerPage  = $ResultsPerPage
            PageCount       = [int]$searchResult.UserArea.NumberOfPages
            TotalJobs       = [int]$searchResult.SearchResultCountAll
            ReturnedCount   = [int]$searchResult.SearchResultCount
            Jobs            = @($searchResult.SearchResultItems)
            RawResponse     = $response
        }
    }
    catch {
        throw "USAJOBS request failed: $($_.Exception.Message)"
    }
    finally {
        $client.Dispose()
    }
}

# ---------------------------------------------------------------------------
# Helpers — null-safe property access
# ---------------------------------------------------------------------------

function Get-Prop {
    param(
        [AllowNull()][object]$Obj,
        [string]$Name,
        [AllowNull()][object]$Default = $null
    )
    if ($null -eq $Obj) { return $Default }
    $p = $Obj.PSObject.Properties[$Name]
    if ($null -eq $p) { return $Default }
    return $p.Value
}

function Normalize-String {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return $null }
    $s = [string]$Value
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    return $s.Trim()
}

function ConvertTo-UtcIsoString {
    param([AllowNull()][object]$Value)
    $s = Normalize-String $Value
    if ($null -eq $s) { return $null }
    try { return ([datetimeoffset]::Parse($s)).UtcDateTime.ToString("o") }
    catch { return $null }
}

function ConvertTo-SafeDecimal {
    param([AllowNull()][object]$Value)
    $s = Normalize-String $Value
    if ($null -eq $s) { return $null }
    try {
        $n = [decimal]$s
        if ($n -gt 0) { return $n }
        return $null
    }
    catch { return $null }
}

# HTML-encode plain text without requiring System.Web
function ConvertTo-HtmlEncoded {
    param([AllowNull()][string]$Text)
    if ($null -eq $Text) { return "" }
    $Text = $Text -replace '&', '&amp;'
    $Text = $Text -replace '<', '&lt;'
    $Text = $Text -replace '>', '&gt;'
    $Text = $Text -replace '"', '&quot;'
    return $Text
}

# Wrap plain text in <p> tags, converting line breaks to HTML
function ConvertTo-HtmlParagraph {
    param([AllowNull()][string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return "" }
    $encoded = ConvertTo-HtmlEncoded $Text.Trim()
    $encoded = $encoded -replace '(\r?\n){2,}', '</p><p>'
    $encoded = $encoded -replace '\r?\n', '<br>'
    return "<p>$encoded</p>"
}

# Build a plain-text summary from a plain-text block (no HTML stripping needed)
function Get-PlainTextSummary {
    param([AllowNull()][string]$Text, [int]$MaxLength = 100)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $s = $Text -replace '[\r\n]+', ' ' -replace '\s{2,}', ' '
    $s = $s.Trim()
    if ($s.Length -le $MaxLength) { return $s }
    return $s.Substring(0, $MaxLength)
}

# ---------------------------------------------------------------------------
# Slug
# ---------------------------------------------------------------------------

function ConvertTo-JobSlug {
    param([string]$Title, [string]$PositionId)
    $slugTitle = $Title.ToLower() `
        -replace '[^a-z0-9\s-]', '' `
        -replace '\s+', '-' `
        -replace '-+', '-'
    $slugId = $PositionId.ToLower() `
        -replace '[^a-z0-9-]', '-' `
        -replace '-+', '-' `
        -replace '^-|-$', ''
    $slug = "$slugTitle-$slugId"
    if ($slug.Length -gt 120) { $slug = $slug.Substring(0, 120).TrimEnd('-') }
    return $slug
}

# ---------------------------------------------------------------------------
# WorkplaceType
# ---------------------------------------------------------------------------

function Get-WorkplaceType {
    param($Descriptor)

    # Check every location — "Anywhere in the U.S." means fully remote
    $locations = @(Get-Prop $Descriptor "PositionLocation")
    foreach ($loc in $locations) {
        $locName = Normalize-String (Get-Prop $loc "LocationName")
        if ($locName -match '(?i)(anywhere|remote|nationwide|no specific location|telework)') {
            return "Remote"
        }
    }

    $locationDisplay = Normalize-String (Get-Prop $Descriptor "PositionLocationDisplay")
    if ($locationDisplay -match '(?i)(anywhere|multiple locations|nationwide)') {
        return "Remote"
    }

    # Check telework eligibility
    $details  = Get-Prop (Get-Prop $Descriptor "UserArea") "Details"
    $telework = Normalize-String (Get-Prop $details "Telework")
    if ($telework -match '(?i)^yes') {
        return "Hybrid"
    }

    return "OnSite"
}

# ---------------------------------------------------------------------------
# EmploymentType — derived from PositionSchedule + PositionOfferingType
# ---------------------------------------------------------------------------

function Get-EmploymentType {
    param($Descriptor)

    $offeringTypes = @(Get-Prop $Descriptor "PositionOfferingType")
    $scheduleTypes = @(Get-Prop $Descriptor "PositionSchedule")

    $offeringName = Normalize-String (Get-Prop ($offeringTypes | Select-Object -First 1) "Name")
    $scheduleName = Normalize-String (Get-Prop ($scheduleTypes | Select-Object -First 1) "Name")

    if ($offeringName -match '(?i)internship|student') { return "Internship" }
    if ($offeringName -match '(?i)\btemporary\b')       { return "Temporary" }
    if ($offeringName -match '(?i)\bterm\b')            { return "Contract" }
    if ($scheduleName -match '(?i)part[\s\-]?time')     { return "PartTime" }
    if ($scheduleName -match '(?i)full[\s\-]?time')     { return "FullTime" }

    return $null    # removed from output when null — API uses its default
}

# ---------------------------------------------------------------------------
# JobCategory — OPM series code → JobCategory enum string
# ---------------------------------------------------------------------------

function Get-JobCategory {
    param($JobCategories)

    $cats = @($JobCategories)
    if ($cats.Count -eq 0) { return $null }

    $codeStr = Normalize-String (Get-Prop ($cats | Select-Object -First 1) "Code")
    $nameStr = Normalize-String (Get-Prop ($cats | Select-Object -First 1) "Name")

    # Series-code mapping (OPM GS occupational groups)
    $code = 0
    if ($null -ne $codeStr -and [int]::TryParse($codeStr, [ref]$code)) {
        $series = [math]::Floor($code / 100)
        switch ($series) {
            1  { return "Research" }
            2  { return "HumanResources" }
            { $_ -in 3, 4 } { return "AdminSecretarial" }
            5  { return "Accounting" }
            6  { return "Healthcare" }
            7  { return "Veterinary" }
            8  { return "Engineering" }
            9  { return "LegalServices" }
            10 { return "MediaJournalism" }
            11 { return "Sales" }
            13 { return "Science" }
            14 { return "AdminSecretarial" }
            15 { return "ITSoftware" }
            17 { return "Education" }
            18 { return "SecurityIntelligence" }
            20 { return "Logistics" }
            21 { return "TransportDistribution" }
            22 { return "ITSoftware" }
        }
    }

    # Name-based fallback (mirrors the approach in Get-RemoteOK.ps1)
    $text = @($nameStr) | Where-Object { $_ } | ForEach-Object { $_ } | Select-Object -First 1
    if ([string]::IsNullOrWhiteSpace($text)) { return "Government" }

    switch -Regex ($text) {
        '(?i)information tech|software|cyber|computer|it management'    { return "ITSoftware" }
        '(?i)engineer(ing)?'                                             { return "Engineering" }
        '(?i)health|medical|nurs|physician|doctor|dental'               { return "Healthcare" }
        '(?i)account|financ|budget|auditor'                             { return "Accounting" }
        '(?i)legal|attorney|law|paralegal|counsel'                      { return "LegalServices" }
        '(?i)human resources|personnel'                                  { return "HumanResources" }
        '(?i)education|teacher|training|instructor'                     { return "Education" }
        '(?i)security|intelligence|investigat|police|law enforcement'   { return "SecurityIntelligence" }
        '(?i)logistic|supply chain|procurement'                         { return "Logistics" }
        '(?i)transport|traffic|air traffic'                             { return "TransportDistribution" }
        '(?i)science|biolog|chemist|physicist'                          { return "Science" }
        '(?i)research|analyst'                                          { return "Research" }
        '(?i)social work|social service|community'                      { return "SocialWork" }
        '(?i)pharma'                                                     { return "Pharmaceuticals" }
        '(?i)aerospace|aviation|space'                                   { return "Aerospace" }
        '(?i)construction|facilities|maintenance'                        { return "BuildingConstruction" }
        default                                                          { return "Government" }
    }
}

# ---------------------------------------------------------------------------
# Description HTML builder
# ---------------------------------------------------------------------------

function Build-Description {
    param($Descriptor)

    $details      = Get-Prop (Get-Prop $Descriptor "UserArea") "Details"
    $jobSummary   = Normalize-String (Get-Prop $details "JobSummary")
    $qualSummary  = Normalize-String (Get-Prop $Descriptor "QualificationSummary")
    $majorDuties  = @(@(Get-Prop $details "MajorDuties" @()) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $keyReqs      = @(@(Get-Prop $details "KeyRequirements" @()) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $education    = Normalize-String (Get-Prop $details "Education")
    $evaluations  = Normalize-String (Get-Prop $details "Evaluations")
    $otherInfo    = Normalize-String (Get-Prop $details "OtherInformation")

    $parts = [System.Collections.Generic.List[string]]::new()

    # Summary / overview — prefer JobSummary, fall back to QualificationSummary
    $summaryText = if ($jobSummary) { $jobSummary } elseif ($qualSummary) { $qualSummary } else { $null }
    if ($summaryText) {
        $parts.Add("<h2>Summary</h2>$(ConvertTo-HtmlParagraph $summaryText)")
    }

    # Major duties
    if ($majorDuties.Count -gt 0) {
        $li = ($majorDuties | ForEach-Object { "<li>$(ConvertTo-HtmlEncoded $_)</li>" }) -join ""
        $parts.Add("<h2>Major Duties</h2><ul>$li</ul>")
    }

    # Qualifications (show separately when a distinct JobSummary was already shown)
    if ($qualSummary -and $jobSummary -and $qualSummary -ne $jobSummary) {
        $parts.Add("<h2>Qualifications</h2>$(ConvertTo-HtmlParagraph $qualSummary)")
    }

    # Key requirements
    if ($keyReqs.Count -gt 0) {
        $li = ($keyReqs | ForEach-Object { "<li>$(ConvertTo-HtmlEncoded $_)</li>" }) -join ""
        $parts.Add("<h2>Key Requirements</h2><ul>$li</ul>")
    }

    # Education
    if ($education) {
        $parts.Add("<h2>Education</h2>$(ConvertTo-HtmlParagraph $education)")
    }

    # Evaluations
    if ($evaluations) {
        $parts.Add("<h2>How You Will Be Evaluated</h2>$(ConvertTo-HtmlParagraph $evaluations)")
    }

    # Other information
    if ($otherInfo) {
        $parts.Add("<h2>Other Information</h2>$(ConvertTo-HtmlParagraph $otherInfo)")
    }

    if ($parts.Count -eq 0) {
        return "<p>Please visit the application URL for full job details.</p>"
    }

    $html = $parts -join ""

    # Truncate to stay within the 20 000-char DB limit (JobConfiguration.HasMaxLength(20000)).
    $maxLen  = 19800
    $trailer = "<p><em>Description truncated. View full details on the source site.</em></p>"
    if ($html.Length -gt $maxLen) {
        # Walk back from the limit to avoid cutting inside an HTML tag.
        $cutAt = $maxLen
        while ($cutAt -gt 0 -and $html[$cutAt] -ne '>') { $cutAt-- }
        $html = $html.Substring(0, $cutAt + 1) + $trailer
    }

    return $html
}

# ---------------------------------------------------------------------------
# Main — fetch 500 jobs (page 1 only) and transform
# ---------------------------------------------------------------------------

Write-Host "Fetching USAJobs (page 1, up to 500 results)..."

try {
    $Data = Get-USAJobsSearchResult `
        -ApiKey $USAJOBS_API_KEY `
        -UserAgentEmail $USAJOBS_USER_AGENT `
        -Page 1 `
        -ResultsPerPage 500
}
catch {
    Write-Warning "USAJobs API request failed — writing empty output and skipping. ($_)"
    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    [System.IO.File]::WriteAllText($OutputPath, "[]", $utf8NoBom)
    exit 0
}

if ($null -eq $Data -or $Data.ReturnedCount -eq 0) {
    Write-Warning "USAJobs API returned no jobs — writing empty output and skipping."
    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    [System.IO.File]::WriteAllText($OutputPath, "[]", $utf8NoBom)
    exit 0
}

Write-Host "Returned: $($Data.ReturnedCount) jobs  |  Total available: $($Data.TotalJobs)  |  Pages: $($Data.PageCount)"

# Deduplicate within the batch by PositionID
$seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

$transformed = foreach ($item in $Data.Jobs) {
    $d = $item.MatchedObjectDescriptor
    if ($null -eq $d) { continue }

    $positionId = Normalize-String (Get-Prop $d "PositionID")
    if ($null -eq $positionId) { continue }

    if (-not $seen.Add($positionId)) {
        Write-Warning "Duplicate PositionID skipped: $positionId"
        continue
    }

    $title = Normalize-String (Get-Prop $d "PositionTitle")
    if ($null -eq $title) { continue }

    # External application URL — prefer ApplyURI[0], fall back to PositionURI
    $applyUris  = @(Get-Prop $d "ApplyURI" @())
    $applyUrl   = Normalize-String ($applyUris | Select-Object -First 1)
    if ($null -eq $applyUrl) {
        $applyUrl = Normalize-String (Get-Prop $d "PositionURI")
    }
    if ($null -eq $applyUrl) { continue }

    # Location
    $locations     = @(Get-Prop $d "PositionLocation" @())
    $firstLocation = $locations | Select-Object -First 1

    $locationText  = Normalize-String (Get-Prop $d "PositionLocationDisplay")
    $locationCity  = Normalize-String (Get-Prop $firstLocation "CityName")
    $locationState = Normalize-String (Get-Prop $firstLocation "CountrySubDivisionCode")

    $locationLat = $null
    $locationLon = $null
    $rawLat = Get-Prop $firstLocation "Latitude"
    $rawLon = Get-Prop $firstLocation "Longitude"
    if ($null -ne $rawLat) { try { $locationLat = [double]$rawLat } catch {} }
    if ($null -ne $rawLon) { try { $locationLon = [double]$rawLon } catch {} }

    # Salary — only use per-annum (PA) figures
    $salaryFrom     = $null
    $salaryTo       = $null
    $salaryCurrency = $null
    $remunerations  = @(Get-Prop $d "PositionRemuneration" @())
    $firstRem       = $remunerations | Select-Object -First 1
    if ($null -ne $firstRem) {
        $rateCode = Normalize-String (Get-Prop $firstRem "RateIntervalCode")
        if ($rateCode -eq "PA") {
            $salaryFrom     = ConvertTo-SafeDecimal (Get-Prop $firstRem "MinimumRange")
            $salaryTo       = ConvertTo-SafeDecimal (Get-Prop $firstRem "MaximumRange")
            $salaryCurrency = "USD"
        }
    }

    # Dates
    $publishedUtc = ConvertTo-UtcIsoString (Get-Prop $d "PublicationStartDate")
    $expiresUtc   = ConvertTo-UtcIsoString (Get-Prop $d "ApplicationCloseDate")

    # Keywords from job categories
    $jobCats = @(Get-Prop $d "JobCategory" @())
    $keywords = if ($jobCats.Count -gt 0) {
        ($jobCats | ForEach-Object { Normalize-String (Get-Prop $_ "Name") } | Where-Object { $_ }) -join ","
    } else { $null }

    # Requirements — KeyRequirements array or Requirements text
    $details  = Get-Prop (Get-Prop $d "UserArea") "Details"
    $keyReqs  = @(@(Get-Prop $details "KeyRequirements" @()) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $requirements = $null
    if ($keyReqs.Count -gt 0) {
        $li = ($keyReqs | ForEach-Object { "<li>$(ConvertTo-HtmlEncoded $_)</li>" }) -join ""
        $requirements = "<ul>$li</ul>"
    } else {
        $requirements = Normalize-String (Get-Prop $details "Requirements")
    }

    # Benefits
    $benefitsText = Normalize-String (Get-Prop $details "Benefits")

    # Employment / workplace
    $empType       = Get-EmploymentType -Descriptor $d
    $workplaceType = Get-WorkplaceType  -Descriptor $d

    $jobObj = [PSCustomObject][ordered]@{
        sourceSite             = "usajobs.gov"
        externalId             = $positionId
        companyName            = Normalize-String (Get-Prop $d "OrganizationName")
        companyLogoUrl         = $COMPANY_LOGO_URL

        title                  = $title
        description            = Build-Description -Descriptor $d
        workplaceType          = $workplaceType
        employmentType         = $empType

        externalApplicationUrl = $applyUrl

        category               = Get-JobCategory -JobCategories $jobCats
        keywords               = $keywords
        regions                = @("NorthAmerica")
        countries              = @("United States")

        summary                = Get-PlainTextSummary -Text (Normalize-String (Get-Prop $details "JobSummary")) -MaxLength 100
        requirements           = $requirements
        benefits               = $benefitsText
        department             = Normalize-String (Get-Prop $d "DepartmentName")

        salaryFrom             = $salaryFrom
        salaryTo               = $salaryTo
        salaryCurrency         = $salaryCurrency

        publishedUtc           = $publishedUtc
        postingExpiresUtc      = $expiresUtc

        locationText           = $locationText
        locationCity           = $locationCity
        locationState          = $locationState
        locationCountry        = "United States"
        locationCountryCode    = "US"
        locationLatitude       = $locationLat
        locationLongitude      = $locationLon

        slug                   = ConvertTo-JobSlug -Title $title -PositionId $positionId
    }

    # employmentType is a non-nullable enum — remove the property when unknown so
    # the API deserializes to its default rather than rejecting the record.
    if ($null -eq $jobObj.employmentType) {
        $jobObj.PSObject.Properties.Remove('employmentType')
    }

    $jobObj
}

if (@($transformed).Count -eq 0) {
    Write-Warning "No valid USAJobs records after transformation — writing empty output and skipping."
    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    [System.IO.File]::WriteAllText($OutputPath, "[]", $utf8NoBom)
    exit 0
}

# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------

$directory = Split-Path -Path $OutputPath -Parent
if (-not [string]::IsNullOrWhiteSpace($directory) -and -not (Test-Path $directory)) {
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
}

$jsonContent = @($transformed) | ConvertTo-Json -Depth 20 -AsArray

try {
    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    [System.IO.File]::WriteAllText($OutputPath, $jsonContent, $utf8NoBom)
}
catch {
    throw "Failed to write output to '$OutputPath': $_"
}

Write-Host "Transformed job count : $(@($transformed).Count)"
Write-Host "Output written to     : $OutputPath"

# ---------------------------------------------------------------------------
# Phase 2: Company logo resolution via usajobs.map
#
# Map file format (plain text, one entry per line):
#   CompanyName=https://example.com/path/to/logo.png   ← has a logo URL
#   CompanyName=                                         ← no logo found yet
#
# Logic:
#   - Read all unique companyName values from the generated JSON
#   - Any name missing from the map → append "name=" (no URL yet)
#   - Names with an empty value → keep default companyLogoUrl, skip
#   - Names with a URL → download the file, rewrite companyLogoUrl to GitHub raw URL
# ---------------------------------------------------------------------------

$mapPath       = Join-Path -Path $PSScriptRoot -ChildPath "usajobs.map"
$logoLocalRoot = Join-Path -Path $PSScriptRoot -ChildPath "usajobs"
$githubRawBase = "https://raw.githubusercontent.com/repasscloud/aethon-software-import-data/refs/heads/main/usajobs"
$utf8NoBom     = [System.Text.UTF8Encoding]::new($false)

# Wikimedia (and other sites) require a descriptive bot User-Agent, not a browser string.
# Format: AppName/version (project-url; contact-email)
$downloadUserAgent = "AethonJobImport/1.0 (https://github.com/repasscloud/aethon-software-import-data; hello@repasscloud.com) PowerShell/$($PSVersionTable.PSVersion)"

# Per-host last-request timestamp — used to enforce a minimum gap between requests
# to the same server (Wikimedia is particularly strict).
$hostLastRequest = @{}

function Invoke-ThrottledDownload {
    <#
    .SYNOPSIS
        Downloads a file with per-host rate limiting, a bot-friendly User-Agent,
        429 retry-with-backoff, and HTML error-page detection.
    #>
    param(
        [string]$Uri,
        [string]$OutFile,
        [string]$UserAgent,
        [int]$InterRequestDelayMs = 1500,   # min gap between requests to the same host
        [int]$MaxRetries           = 4
    )

    $uriObj   = [System.Uri]$Uri
    $hostname = $uriObj.Host

    # Enforce per-host delay
    if ($hostLastRequest.ContainsKey($hostname)) {
        $elapsed = ([datetime]::UtcNow - $hostLastRequest[$hostname]).TotalMilliseconds
        if ($elapsed -lt $InterRequestDelayMs) {
            Start-Sleep -Milliseconds ([int]($InterRequestDelayMs - $elapsed))
        }
    }

    $tmpFile = [System.IO.Path]::GetTempFileName()

    try {
        for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
            $hostLastRequest[$hostname] = [datetime]::UtcNow

            try {
                $response = Invoke-WebRequest `
                    -Uri            $Uri `
                    -OutFile        $tmpFile `
                    -UseBasicParsing `
                    -TimeoutSec     45 `
                    -Headers        @{ 'User-Agent' = $UserAgent } `
                    -SkipHttpErrorCheck `
                    -PassThru

                $status = [int]$response.StatusCode

                if ($status -eq 429) {
                    # Respect Retry-After if present; otherwise use exponential backoff
                    $retryAfter = 30 * $attempt
                    $raHeader   = $response.Headers['Retry-After']
                    if ($raHeader) {
                        $parsed = 0
                        if ([int]::TryParse([string]$raHeader, [ref]$parsed) -and $parsed -gt 0) {
                            $retryAfter = $parsed
                        }
                    }
                    Write-Host "    429 rate-limited — waiting ${retryAfter}s (attempt $attempt/$MaxRetries)..."
                    if (Test-Path $tmpFile) { Remove-Item $tmpFile -Force }
                    $tmpFile = [System.IO.Path]::GetTempFileName()
                    Start-Sleep -Seconds $retryAfter
                    continue
                }

                if ($status -ge 400) {
                    throw "HTTP $status from $Uri"
                }

                # Verify the saved file is not an HTML error page
                $head = [System.IO.File]::ReadAllText($tmpFile, [System.Text.Encoding]::UTF8) |
                    Select-String -Pattern '^\s*<(!DOCTYPE|html)' -CaseSensitive:$false
                if ($head) {
                    throw "Response body is an HTML page (server-side error), not an image."
                }

                # All good — move temp file to final destination
                if (Test-Path $OutFile) { Remove-Item $OutFile -Force }
                [System.IO.File]::Move($tmpFile, $OutFile)
                $tmpFile = $null
                return $true
            }
            catch {
                if ($tmpFile -and (Test-Path $tmpFile)) {
                    Remove-Item $tmpFile -Force
                    $tmpFile = [System.IO.Path]::GetTempFileName()
                }
                if ($attempt -ge $MaxRetries) { throw }
                $backoff = [math]::Pow(2, $attempt) * 3
                Write-Host "    Error on attempt $attempt — waiting ${backoff}s before retry..."
                Start-Sleep -Seconds $backoff
            }
        }
    }
    finally {
        if ($tmpFile -and (Test-Path $tmpFile)) { Remove-Item $tmpFile -Force }
    }

    return $false
}

Write-Host ""
Write-Host "--- Phase 2: company logo resolution ---"

# Load the map file into an ordered dictionary (preserves insertion order for writes)
$map = [System.Collections.Specialized.OrderedDictionary]::new()

if (Test-Path $mapPath) {
    foreach ($line in [System.IO.File]::ReadAllLines($mapPath, [System.Text.Encoding]::UTF8)) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $eqIdx = $line.IndexOf('=')
        if ($eqIdx -lt 0) { continue }
        $k = $line.Substring(0, $eqIdx)
        $v = $line.Substring($eqIdx + 1)
        $map[$k] = $v
    }
    Write-Host "Loaded $($map.Count) entries from usajobs.map"
} else {
    Write-Host "usajobs.map not found — will create it."
}

# Re-read the JSON we just wrote so we work from the serialised form
$importData = Get-Content -Path $OutputPath -Raw -Encoding UTF8 | ConvertFrom-Json

# Collect unique company names from the output
$uniqueCompanies = @($importData | ForEach-Object { $_.companyName } | Sort-Object -Unique)
Write-Host "Unique company names in output: $($uniqueCompanies.Count)"

# Add any company not yet in the map (empty value = not yet researched)
$mapChanged = $false
foreach ($name in $uniqueCompanies) {
    if (-not $map.Contains($name)) {
        $map[$name] = ""
        $mapChanged = $true
        Write-Host "  MAP +  '$name' (no logo yet)"
    }
}

# Save map if new entries were added
if ($mapChanged) {
    $lines = foreach ($k in $map.Keys) { "$k=$($map[$k])" }
    [System.IO.File]::WriteAllLines($mapPath, $lines, $utf8NoBom)
    Write-Host "usajobs.map updated."
}

# Process each company that has a logo URL in the map
$anyJsonUpdated = $false

foreach ($name in $uniqueCompanies) {
    $logoUrl = $map[$name]
    if ([string]::IsNullOrWhiteSpace($logoUrl)) {
        # No URL in map — default logo stays, nothing to do
        continue
    }

    # Derive the local path from the URL (strip scheme + domain, keep path)
    try { $uriObj = [System.Uri]$logoUrl } catch {
        Write-Warning "  SKIP  '$name' — invalid URL: $logoUrl"
        continue
    }
    $relativePath = $uriObj.AbsolutePath.TrimStart('/')   # e.g. wikipedia/commons/a/a3/file.svg
    $localPath    = Join-Path -Path $logoLocalRoot -ChildPath $relativePath
    $githubUrl    = "$githubRawBase/$relativePath"

    # Ensure the sub-directory exists
    $localDir = Split-Path -Path $localPath -Parent
    if (-not (Test-Path $localDir)) {
        New-Item -ItemType Directory -Path $localDir -Force | Out-Null
    }

    # Download only if not already cached
    if (-not (Test-Path $localPath)) {
        try {
            $ok = Invoke-ThrottledDownload `
                -Uri                 $logoUrl `
                -OutFile             $localPath `
                -UserAgent           $downloadUserAgent `
                -InterRequestDelayMs 1500 `
                -MaxRetries          4
            if ($ok) {
                Write-Host "  DL    '$name' → $relativePath"
            } else {
                Write-Warning "  FAIL  '$name' — download returned false after retries"
                continue
            }
        }
        catch {
            Write-Warning "  FAIL  '$name' — $_"
            continue
        }
    } else {
        Write-Host "  CACHE '$name' → $relativePath"
    }

    # Update all matching records in the JSON (a company may have multiple postings)
    foreach ($record in $importData) {
        if ($record.companyName -eq $name) {
            $record.companyLogoUrl = $githubUrl
            $anyJsonUpdated = $true
        }
    }
}

# Write the updated JSON back if any logo URLs changed
if ($anyJsonUpdated) {
    $updatedJson = @($importData) | ConvertTo-Json -Depth 20 -AsArray
    [System.IO.File]::WriteAllText($OutputPath, $updatedJson, $utf8NoBom)
    Write-Host "JSON updated with resolved logo URLs."
} else {
    Write-Host "No logo URLs required updating."
}

Write-Host "--- Phase 2 complete ---"

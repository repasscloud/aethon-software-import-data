# Aethon Import Script Generator

This directory contains PowerShell scripts that fetch jobs from external APIs and
transform them into the Aethon `ImportJobDto` JSON format for bulk import via
`POST /api/v1/import/jobs/bulk`.

When asked to create a new import script for a data source, follow every rule in
this file exactly.

---

## Output file contract

Every script must produce a **flat JSON array** at the root level — `[{...},{...}]`.
Each element maps directly to one `ImportJobDto`.

> **Critical:** Never use `ConvertTo-Json -InputObject $variable`. Always pipe:
> `$variable | ConvertTo-Json -Depth 20 -AsArray`
> Using `-InputObject` with an array causes PowerShell to double-wrap it → `[[...]]`,
> which makes every record fail with a parse error on the API.

---

## ImportJobDto field reference

These are the exact JSON property names and their types. All names are camelCase
(the API's `JsonStringEnumConverter` is case-insensitive for enums but the property
names must be camelCase).

| JSON property | Type | Required | Notes |
|---|---|---|---|
| `sourceSite` | string | yes | e.g. `"linkedin.com"` |
| `externalId` | string | yes | ID from source system |
| `companyName` | string | yes | |
| `companyLogoUrl` | string? | no | Full URL or null |
| `title` | string | yes | |
| `description` | string | yes | HTML or plain text, max 20 000 chars |
| `workplaceType` | WorkplaceType enum | yes | See enum values below |
| `employmentType` | EmploymentType? enum | no | Omit or null when unknown |
| `externalApplicationUrl` | string | yes | Deep-link to apply |
| `category` | JobCategory? enum | no | See enum values below |
| `keywords` | string? | no | Comma-separated string or null |
| `regions` | JobRegion[] | yes | Empty array `[]` is valid |
| `countries` | string[] | yes | Country names from the canonical list |
| `summary` | string? | no | Plain text, ≤ 100 chars |
| `requirements` | string? | no | HTML or plain text |
| `benefits` | string? | no | Plain text |
| `department` | string? | no | |
| `salaryFrom` | decimal? | no | Annual figures only; null otherwise |
| `salaryTo` | decimal? | no | Annual figures only; null otherwise |
| `salaryCurrency` | CurrencyCode? enum | no | See enum values below; null for unknowns |
| `publishedUtc` | DateTime? ISO 8601 | no | UTC, `"o"` format |
| `postingExpiresUtc` | DateTime? ISO 8601 | no | UTC, `"o"` format |
| `locationText` | string? | no | Raw location string from source |
| `locationCity` | string? | no | |
| `locationState` | string? | no | State/province code |
| `locationCountry` | string? | no | Canonical country name |
| `locationCountryCode` | string? | no | ISO 3166-1 alpha-2 |
| `locationLatitude` | double? | no | |
| `locationLongitude` | double? | no | |
| `slug` | string? | no | URL-safe slug, max 120 chars |

---

## Enum allowed values

### WorkplaceType (required)
`Remote` | `Hybrid` | `OnSite`

### EmploymentType (nullable — omit the property when unknown)
`FullTime` | `PartTime` | `Contract` | `Temporary` | `Casual` | `Internship`

### JobRegion
`Africa` | `Asia` | `Europe` | `LatinAmerica` | `MiddleEast` | `NorthAmerica` | `Oceania` | `Worldwide`

### JobCategory
`Accounting` | `AdminSecretarial` | `Banking` | `ExecutiveManagement` | `FinanceInsurance` | `HumanResources` | `LegalServices` | `Recruitment` | `AdvertisingPR` | `Arts` | `Design` | `Marketing` | `MediaJournalism` | `PublicRelations` | `ITSoftware` | `TelecommunicationsISP` | `Aerospace` | `Engineering` | `Research` | `Science` | `SecurityIntelligence` | `AgricultureFishingForestry` | `MiningResources` | `UtilitiesEnergy` | `Veterinary` | `BuildingConstruction` | `PropertyRealEstate` | `Automobile` | `Logistics` | `Manufacturing` | `TransportDistribution` | `Education` | `Government` | `GraduateRoles` | `SocialWork` | `Charity` | `Healthcare` | `Pharmaceuticals` | `Catering` | `CustomerService` | `FoodBeverage` | `Hospitality` | `Tourism` | `Retail` | `Sales` | `PartTimeTemp` | `SportRecreation` | `Other`

### CurrencyCode (nullable — use null for any code not in this list)
`AUD` | `USD` | `EUR` | `GBP` | `NZD` | `CAD` | `CHF` | `JPY` | `CNY` | `HKD` | `SEK` | `KRW` | `SGD` | `NOK` | `MXN` | `INR` | `RUB` | `ZAR` | `TRY` | `BRL` | `RON`

---

## Required helper functions

Every script must include these helpers (copy verbatim, they are well-tested):

### Null-safe property accessor
```powershell
function Get-Prop {
    param([AllowNull()][object]$Obj, [string]$Name, [AllowNull()][object]$Default = $null)
    if ($null -eq $Obj) { return $Default }
    $p = $Obj.PSObject.Properties[$Name]
    if ($null -eq $p) { return $Default }
    return $p.Value
}
```

### String normalisation
```powershell
function Normalize-String {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return $null }
    $s = [string]$Value
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    return $s.Trim()
}
```

### UTC ISO-8601 date
```powershell
function ConvertTo-UtcIsoString {
    param([AllowNull()][object]$Value)
    $s = Normalize-String $Value
    if ($null -eq $s) { return $null }
    try { return ([datetimeoffset]::Parse($s)).UtcDateTime.ToString("o") }
    catch { return $null }
}
```

### Safe decimal (for salary fields)
```powershell
function ConvertTo-SafeDecimal {
    param([AllowNull()][object]$Value)
    $s = Normalize-String $Value
    if ($null -eq $s) { return $null }
    try { $n = [decimal]$s; if ($n -gt 0) { return $n }; return $null }
    catch { return $null }
}
```

### Currency code guard — MUST be used for any currency value from an external API
```powershell
$ValidCurrencyCodes = @(
    "AUD","USD","EUR","GBP","NZD","CAD","CHF","JPY","CNY","HKD",
    "SEK","KRW","SGD","NOK","MXN","INR","RUB","ZAR","TRY","BRL","RON"
)
function Normalize-CurrencyCode {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return $null }
    $code = ([string]$Value).Trim().ToUpperInvariant()
    if ([string]::IsNullOrWhiteSpace($code)) { return $null }
    if ($ValidCurrencyCodes -contains $code) { return $code }
    return $null   # unknown currency → omit rather than crash the record
}
```

### URL slug
```powershell
function ConvertTo-JobSlug {
    param([string]$Title, [string]$Id)
    $slugTitle = $Title.ToLower() -replace '[^a-z0-9\s-]','' -replace '\s+','-' -replace '-+','-'
    $suffix = $Id -replace '^.*/','' -replace '[^a-z0-9-]','-' -replace '-+','-' -replace '^-|-$',''
    if ($suffix.Length -gt 20) { $suffix = $suffix.Substring(0,20).TrimEnd('-') }
    $slug = "$slugTitle-$suffix"
    if ($slug.Length -gt 120) { $slug = $slug.Substring(0,120).TrimEnd('-') }
    return $slug.ToLower()
}
```

---

## File naming convention

| Source | Output filename |
|---|---|
| `Get-Jobicy.ps1` | `jobicy-import.json` |
| `Get-RemoteOK.ps1` | `remoteok-import.json` |
| `Get-USAJobs.ps1` | `usajobs-import.json` |
| `Get-Jooble.ps1` | `jooble_{country}-import.json` (one file per country) |
| New single-feed | `Get-{SourceName}.ps1` → `{sourcename}-import.json` |
| New per-country | `Get-{SourceName}.ps1` → `{sourcename}_{country}-import.json` |

---

## Script structure (template)

```powershell
[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path -Path $PSScriptRoot -ChildPath "{sourcename}-import.json")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# API credentials loaded from environment variables — never hard-coded
$API_KEY = $env:SOURCENAME_API_KEY

if ([string]::IsNullOrWhiteSpace($API_KEY)) {
    Write-Warning "SOURCENAME_API_KEY is missing — writing empty output and skipping."
    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    [System.IO.File]::WriteAllText($OutputPath, "[]", $utf8NoBom)
    exit 0
}

# ... helper functions (Get-Prop, Normalize-String, etc.) ...

# Fetch
try {
    # ... API call ...
}
catch {
    Write-Warning "API request failed — writing empty output and skipping. ($_)"
    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    [System.IO.File]::WriteAllText($OutputPath, "[]", $utf8NoBom)
    exit 0
}

# Transform
$transformed = foreach ($job in $rawJobs) {
    # ... map fields to ImportJobDto shape ...
    [PSCustomObject][ordered]@{
        sourceSite             = "{sourcename}"
        externalId             = ...
        # etc.
    }
}

# Write — MUST pipe, never use -InputObject
$directory = Split-Path -Path $OutputPath -Parent
if (-not [string]::IsNullOrWhiteSpace($directory) -and -not (Test-Path $directory)) {
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
}

$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
$jsonContent = @($transformed) | ConvertTo-Json -Depth 20 -AsArray
[System.IO.File]::WriteAllText($OutputPath, $jsonContent, $utf8NoBom)

Write-Host "Transformed job count : $(@($transformed).Count)"
Write-Host "Output written to     : $OutputPath"
```

---

## Rules checklist

Before finishing a new script, verify:

- [ ] Env var is checked and script exits cleanly (writes `[]`) if missing
- [ ] API fetch is wrapped in try/catch that writes `[]` and exits 0 on failure
- [ ] Every record that lacks `externalId`, `title`, or `externalApplicationUrl` is skipped with `continue`
- [ ] Duplicate `externalId` values are deduplicated with a `HashSet`
- [ ] `salaryCurrency` always goes through `Normalize-CurrencyCode`
- [ ] `employmentType` is omitted from the object (not set to null) when unknown
- [ ] `description` is capped at 20 000 characters
- [ ] `summary` is capped at 100 characters (plain text, no HTML)
- [ ] `slug` is max 120 characters, lowercase alphanumeric + hyphens only
- [ ] Final JSON write uses `@($transformed) | ConvertTo-Json -Depth 20 -AsArray` (pipe, not `-InputObject`)
- [ ] File is written with `[System.IO.File]::WriteAllText(..., $utf8NoBom)` (no BOM, no trailing newline added by Set-Content)
- [ ] `regions` and `countries` are always arrays (even when empty)
- [ ] All API credentials come from `$env:VARIABLE_NAME`, never from hard-coded strings

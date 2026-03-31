[CmdletBinding()]
param(
    [string]$ApiUrl = "https://jobicy.com/api/v2/remote-jobs",
    [string]$OutputPath = (Join-Path -Path $PSScriptRoot -ChildPath "jobicy-import.json")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Canonical region -> countries map from your platform rules.
$RegionCountryMap = @{
    Africa = @(
        "Algeria","Angola","Benin","Botswana","Burkina Faso","Burundi","Cameroon","Cape Verde",
        "Central African Republic","Chad","Comoros","Congo","DR Congo","Djibouti","Egypt",
        "Equatorial Guinea","Eritrea","Eswatini","Ethiopia","Gabon","Gambia","Ghana","Guinea",
        "Guinea-Bissau","Ivory Coast","Kenya","Lesotho","Liberia","Libya","Madagascar","Malawi",
        "Mali","Mauritania","Mauritius","Morocco","Mozambique","Namibia","Niger","Nigeria",
        "Rwanda","São Tomé and Príncipe","Senegal","Seychelles","Sierra Leone","Somalia",
        "South Africa","South Sudan","Sudan","Tanzania","Togo","Tunisia","Uganda","Zambia","Zimbabwe"
    )
    Asia = @(
        "Afghanistan","Armenia","Azerbaijan","Bahrain","Bangladesh","Bhutan","Brunei","Cambodia",
        "China","Cyprus","Georgia","India","Indonesia","Iran","Iraq","Israel","Japan","Jordan",
        "Kazakhstan","Kuwait","Kyrgyzstan","Laos","Lebanon","Malaysia","Maldives","Mongolia",
        "Myanmar","Nepal","North Korea","Oman","Pakistan","Palestine","Philippines","Qatar",
        "Saudi Arabia","Singapore","South Korea","Sri Lanka","Syria","Taiwan","Tajikistan",
        "Thailand","Timor-Leste","Turkey","Turkmenistan","United Arab Emirates","Uzbekistan",
        "Vietnam","Yemen"
    )
    Europe = @(
        "Albania","Andorra","Austria","Belarus","Belgium","Bosnia and Herzegovina","Bulgaria",
        "Croatia","Cyprus","Czech Republic","Denmark","Estonia","Finland","France","Germany",
        "Greece","Hungary","Iceland","Ireland","Italy","Kosovo","Latvia","Liechtenstein",
        "Lithuania","Luxembourg","Malta","Moldova","Monaco","Montenegro","Netherlands",
        "North Macedonia","Norway","Poland","Portugal","Romania","Russia","San Marino","Serbia",
        "Slovakia","Slovenia","Spain","Sweden","Switzerland","Ukraine","United Kingdom","Vatican City"
    )
    LatinAmerica = @(
        "Argentina","Belize","Bolivia","Brazil","Chile","Colombia","Costa Rica","Cuba",
        "Dominican Republic","Ecuador","El Salvador","Guatemala","Guyana","Haiti","Honduras",
        "Jamaica","Mexico","Nicaragua","Panama","Paraguay","Peru","Puerto Rico","Suriname",
        "Trinidad and Tobago","Uruguay","Venezuela"
    )
    MiddleEast = @(
        "Bahrain","Egypt","Iran","Iraq","Israel","Jordan","Kuwait","Lebanon","Oman","Palestine",
        "Qatar","Saudi Arabia","Syria","Turkey","United Arab Emirates","Yemen"
    )
    NorthAmerica = @(
        "Canada","Mexico","United States"
    )
    Oceania = @(
        "Australia","Fiji","Kiribati","Marshall Islands","Micronesia","Nauru","New Zealand",
        "Palau","Papua New Guinea","Samoa","Solomon Islands","Tonga","Tuvalu","Vanuatu"
    )
    Worldwide = @()
}

# Build a reverse lookup: country -> one or more regions.
$CountryToRegions = @{}
foreach ($region in $RegionCountryMap.Keys) {
    foreach ($country in $RegionCountryMap[$region]) {
        if (-not $CountryToRegions.ContainsKey($country)) {
            $CountryToRegions[$country] = New-Object System.Collections.Generic.List[string]
        }

        if (-not $CountryToRegions[$country].Contains($region)) {
            [void]$CountryToRegions[$country].Add($region)
        }
    }
}

# Add aliases/synonyms used by source feeds.
$CountryAliases = @{
    "UK"      = "United Kingdom"
    "USA"     = "United States"
    "US"      = "United States"
    "U.S."    = "United States"
    "U.S.A."  = "United States"
}

function Get-ObjectPropertyValue {
    param(
        [Parameter(Mandatory)]
        [object]$InputObject,

        [Parameter(Mandatory)]
        [string]$PropertyName,

        [AllowNull()]
        [object]$DefaultValue = $null
    )

    if ($null -eq $InputObject) {
        return $DefaultValue
    }

    $property = $InputObject.PSObject.Properties[$PropertyName]
    if ($null -eq $property) {
        return $DefaultValue
    }

    return $property.Value
}

function Normalize-ToArray {
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return @()
    }

    if ($Value -is [System.Array]) {
        return @($Value)
    }

    return @($Value)
}

function Normalize-CountryName {
    param(
        [AllowNull()]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    $trimmed = $Value.Trim()

    if ($CountryAliases.ContainsKey($trimmed)) {
        return $CountryAliases[$trimmed]
    }

    return $trimmed
}

function Convert-EmploymentType {
    param(
        [AllowNull()]
        [object]$JobType
    )

    $items = @(Normalize-ToArray -Value $JobType)
    $value = if ($items.Count -gt 0) { [string]$items[0] } else { $null }

    if ([string]::IsNullOrWhiteSpace($value)) {
        return $null
    }

    switch ($value.Trim()) {
        "Full-Time"  { return "FullTime" }
        "Part-Time"  { return "PartTime" }
        "Contract"   { return "Contract" }
        "Temporary"  { return "Temporary" }
        "Casual"     { return "Casual" }
        "Internship" { return "Internship" }
        default      { return $null }
    }
}

function Get-FirstIndustryValue {
    param(
        [AllowNull()]
        [object]$Industry
    )

    $items = @(Normalize-ToArray -Value $Industry)
    if ($items.Count -eq 0) {
        return $null
    }

    return [string]$items[0]
}

function Convert-IndustryToCategory {
    param(
        [AllowNull()]
        [string]$Industry
    )

    if ([string]::IsNullOrWhiteSpace($Industry)) {
        return "Other"
    }

    $value = $Industry.Trim()

    switch -Regex ($value) {
        'Software|Developer|Development|Programming|IT|Technology|Tech|DevOps|QA|Data|Security' { return "ITSoftware" }
        'Engineering|Engineer'                                                                    { return "Engineering" }
        'Design|UX|UI|Graphic|Product Design'                                                     { return "Design" }
        'Marketing|SEO|Content|Brand|Growth|Digital Marketing|Social Media'                       { return "Marketing" }
        'Sales|Account Executive|Business Development'                                             { return "Sales" }
        'Customer Success|Customer Service|Support|Technical Support'                              { return "CustomerService" }
        'Human Resources|HR|Talent|People'                                                         { return "HumanResources" }
        'Finance|Accounting|Insurance|Bookkeep'                                                    { return "FinanceInsurance" }
        'Banking'                                                                                  { return "Banking" }
        'Education|Teaching|Training'                                                              { return "Education" }
        'Healthcare|Medical|Nursing|Clinical'                                                      { return "Healthcare" }
        'Legal'                                                                                    { return "LegalServices" }
        'Recruitment|Recruiter|Staffing'                                                           { return "Recruitment" }
        'Research'                                                                                 { return "Research" }
        'Science'                                                                                  { return "Science" }
        'Media|Journalism|Editorial|Publishing'                                                    { return "MediaJournalism" }
        'Public Relations|PR'                                                                      { return "PublicRelations" }
        'Advertising'                                                                              { return "AdvertisingPR" }
        'Administration|Admin|Secretarial|Office'                                                  { return "AdminSecretarial" }
        'Hospitality|Hotel|Restaurant'                                                             { return "Hospitality" }
        'Catering'                                                                                 { return "Catering" }
        'Retail|Ecommerce|E-commerce'                                                              { return "Retail" }
        'Logistics|Supply Chain|Warehouse'                                                         { return "Logistics" }
        'Transport|Distribution|Driver'                                                            { return "TransportDistribution" }
        'Manufacturing|Production'                                                                 { return "Manufacturing" }
        'Construction|Building'                                                                    { return "BuildingConstruction" }
        'Property|Real Estate'                                                                     { return "PropertyRealEstate" }
        'Pharma|Pharmaceutical'                                                                    { return "Pharmaceuticals" }
        'Telecommunications|Telecom|ISP'                                                           { return "TelecommunicationsISP" }
        'Government|Public Sector'                                                                 { return "Government" }
        'Executive|Leadership|Management|Director|VP|Chief'                                        { return "ExecutiveManagement" }
        'Graduate|Entry Level'                                                                     { return "GraduateRoles" }
        'Part[- ]?Time|Temporary'                                                                  { return "PartTimeTemp" }
        'Tourism|Travel'                                                                           { return "Tourism" }
        'Utilities|Energy'                                                                         { return "UtilitiesEnergy" }
        'Agriculture|Fishing|Forestry'                                                             { return "AgricultureFishingForestry" }
        'Mining|Resources'                                                                         { return "MiningResources" }
        'Sport|Recreation|Fitness'                                                                 { return "SportRecreation" }
        'Arts|Creative'                                                                            { return "Arts" }
        'Charity|Nonprofit|Non-profit|NGO'                                                         { return "Charity" }
        'Food|Beverage'                                                                            { return "FoodBeverage" }
        'Automotive|Automobile'                                                                    { return "Automobile" }
        'Aerospace|Aviation'                                                                       { return "Aerospace" }
        'Veterinary'                                                                               { return "Veterinary" }
        'Social Work|Community'                                                                    { return "SocialWork" }
        default                                                                                    { return "Other" }
    }
}

function Get-GeoParts {
    param(
        [AllowNull()]
        [string]$JobGeo
    )

    if ([string]::IsNullOrWhiteSpace($JobGeo)) {
        return @()
    }

    @(
        $JobGeo -split "," |
        ForEach-Object { $_.Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )
}

function Convert-JobGeoToCountries {
    param(
        [AllowNull()]
        [string]$JobGeo
    )

    $parts = Get-GeoParts -JobGeo $JobGeo
    $countries = New-Object System.Collections.Generic.List[string]

    foreach ($part in $parts) {
        $normalized = Normalize-CountryName -Value $part
        if ($null -ne $normalized -and $CountryToRegions.ContainsKey($normalized)) {
            if (-not $countries.Contains($normalized)) {
                [void]$countries.Add($normalized)
            }
        }
    }

    return @($countries)
}

function Convert-JobGeoToRegions {
    param(
        [AllowNull()]
        [string]$JobGeo
    )

    $parts = Get-GeoParts -JobGeo $JobGeo
    $regions = New-Object System.Collections.Generic.List[string]

    foreach ($part in $parts) {
        $normalized = Normalize-CountryName -Value $part

        switch ($normalized) {
            "Worldwide" {
                if (-not $regions.Contains("Worldwide")) { [void]$regions.Add("Worldwide") }
                continue
            }
            "Global" {
                if (-not $regions.Contains("Worldwide")) { [void]$regions.Add("Worldwide") }
                continue
            }
            "EMEA" {
                foreach ($r in @("Europe", "MiddleEast", "Africa")) {
                    if (-not $regions.Contains($r)) { [void]$regions.Add($r) }
                }
                continue
            }
            "APAC" {
                foreach ($r in @("Asia", "Oceania")) {
                    if (-not $regions.Contains($r)) { [void]$regions.Add($r) }
                }
                continue
            }
            "Europe" {
                if (-not $regions.Contains("Europe")) { [void]$regions.Add("Europe") }
                continue
            }
            "Asia" {
                if (-not $regions.Contains("Asia")) { [void]$regions.Add("Asia") }
                continue
            }
            "Africa" {
                if (-not $regions.Contains("Africa")) { [void]$regions.Add("Africa") }
                continue
            }
            "Middle East" {
                if (-not $regions.Contains("MiddleEast")) { [void]$regions.Add("MiddleEast") }
                continue
            }
            "North America" {
                if (-not $regions.Contains("NorthAmerica")) { [void]$regions.Add("NorthAmerica") }
                continue
            }
            "Latin America" {
                if (-not $regions.Contains("LatinAmerica")) { [void]$regions.Add("LatinAmerica") }
                continue
            }
            "Oceania" {
                if (-not $regions.Contains("Oceania")) { [void]$regions.Add("Oceania") }
                continue
            }
        }

        if ($null -ne $normalized -and $CountryToRegions.ContainsKey($normalized)) {
            foreach ($region in $CountryToRegions[$normalized]) {
                if (-not $regions.Contains($region)) {
                    [void]$regions.Add($region)
                }
            }
        }
    }

    return @($regions)
}

function Convert-ToUtcIsoString {
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return $null
    }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    try {
        return ([datetimeoffset]::Parse($text)).UtcDateTime.ToString("o")
    }
    catch {
        return $null
    }
}

function Convert-SafeDecimal {
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return $null
    }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    try {
        return [decimal]$text
    }
    catch {
        return $null
    }
}

Write-Host "Fetching Jobicy jobs from $ApiUrl"
try {
    $data = Invoke-RestMethod -Uri $ApiUrl -Method Get
}
catch {
    throw "Failed to fetch jobs from '$ApiUrl': $_"
}

if ($null -eq $data) {
    throw "The API returned no data."
}

$jobs = @(Get-ObjectPropertyValue -InputObject $data -PropertyName "jobs" -DefaultValue @())

if ($jobs.Count -eq 0) {
    throw "The API response did not contain any jobs."
}

$transformed = foreach ($job in $jobs) {
    $jobGeo = Get-ObjectPropertyValue -InputObject $job -PropertyName "jobGeo"
    $jobIndustry = Get-ObjectPropertyValue -InputObject $job -PropertyName "jobIndustry"
    $jobType = Get-ObjectPropertyValue -InputObject $job -PropertyName "jobType"

    $countries = @(Convert-JobGeoToCountries -JobGeo $jobGeo)
    $regions = @(Convert-JobGeoToRegions -JobGeo $jobGeo)
    $industry = Get-FirstIndustryValue -Industry $jobIndustry
    $employmentType = Convert-EmploymentType -JobType $jobType
    $category = Convert-IndustryToCategory -Industry $industry

    [PSCustomObject]@{
        sourceSite             = "jobicy.com"
        externalId             = [string](Get-ObjectPropertyValue -InputObject $job -PropertyName "id")
        companyName            = Get-ObjectPropertyValue -InputObject $job -PropertyName "companyName"
        companyLogoUrl         = Get-ObjectPropertyValue -InputObject $job -PropertyName "companyLogo"

        title                  = Get-ObjectPropertyValue -InputObject $job -PropertyName "jobTitle"
        description            = Get-ObjectPropertyValue -InputObject $job -PropertyName "jobDescription"
        workplaceType          = "Remote"
        employmentType         = $employmentType

        externalApplicationUrl = Get-ObjectPropertyValue -InputObject $job -PropertyName "url"

        category               = $category
        keywords               = $null
        regions                = @($regions)
        countries              = @($countries)

        summary                = Get-ObjectPropertyValue -InputObject $job -PropertyName "jobExcerpt"
        requirements           = $null
        benefits               = $null
        department             = $null

        salaryFrom             = Convert-SafeDecimal -Value (Get-ObjectPropertyValue -InputObject $job -PropertyName "salaryMin")
        salaryTo               = Convert-SafeDecimal -Value (Get-ObjectPropertyValue -InputObject $job -PropertyName "salaryMax")
        salaryCurrency         = Get-ObjectPropertyValue -InputObject $job -PropertyName "salaryCurrency"

        publishedUtc           = Convert-ToUtcIsoString -Value (Get-ObjectPropertyValue -InputObject $job -PropertyName "pubDate")
        postingExpiresUtc      = $null

        locationText           = $jobGeo
        locationCity           = $null
        locationState          = $null
        locationCountry        = if ($countries.Count -eq 1) { $countries[0] } else { $null }
        locationCountryCode    = $null
        locationLatitude       = $null
        locationLongitude      = $null

        slug                   = Get-ObjectPropertyValue -InputObject $job -PropertyName "jobSlug"
    }
}

$directory = Split-Path -Path $OutputPath -Parent
if (-not [string]::IsNullOrWhiteSpace($directory) -and -not (Test-Path $directory)) {
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
}

$jsonContent = @($transformed) | ConvertTo-Json -Depth 20
try {
    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    [System.IO.File]::WriteAllText($OutputPath, $jsonContent, $utf8NoBom)
}
catch {
    throw "Failed to write output to '$OutputPath': $_"
}

Write-Host "Saved transformed import JSON to: $OutputPath"
Write-Host "Transformed job count: $(@($transformed).Count)"

# Download company logos and save to local paths, while preserving directory structure
if (-not (Test-Path -LiteralPath $OutputPath)) {
    throw "Input file not found: $OutputPath"
}

# Read and parse the JSON file
$jobs = Get-Content -LiteralPath $OutputPath -Raw | ConvertFrom-Json

if ($null -eq $jobs) {
    throw "No job data found in: $OutputPath"
}

foreach ($job in $jobs) {
    # Skip items without a usable logo URL
    if ([string]::IsNullOrWhiteSpace($job.companyLogoUrl)) {
        Write-Warning "Skipping item with externalId '$($job.externalId)' because companyLogoUrl is empty."
        continue
    }

    try {
        $logoUri = [System.Uri]$job.companyLogoUrl
    }
    catch {
        Write-Warning "Skipping invalid URL for externalId '$($job.externalId)': $($job.companyLogoUrl)"
        continue
    }

    # Build a relative file path from the URL:
    # Example:
    # https://jobicy.com/data/server-nyc0409/.../file.jpg
    # =>
    # jobicy.com/data/server-nyc0409/.../file.jpg
    $relativePath = Join-Path -Path $logoUri.Host -ChildPath ($logoUri.AbsolutePath.TrimStart('/') -replace '/', [IO.Path]::DirectorySeparatorChar)

    # Final destination:
    # Join-Path -Path $PSScriptRoot -ChildPath "jobicy.com/data/server-nyc0409/..."
    $destinationPath = Join-Path -Path $PSScriptRoot -ChildPath $relativePath
    $destinationFolder = Split-Path -Path $destinationPath -Parent

    # Ensure destination folder exists
    if (-not (Test-Path -LiteralPath $destinationFolder)) {
        New-Item -Path $destinationFolder -ItemType Directory -Force | Out-Null
    }

    # Download the file only if it doesn't already exist
    if (-not (Test-Path -LiteralPath $destinationPath)) {
        Write-Host "Downloading: $($job.companyLogoUrl)"
        Write-Host "Saving to : $destinationPath"

        Invoke-WebRequest `
            -Uri $job.companyLogoUrl `
            -OutFile $destinationPath
    }
    else {
        Write-Host "Already exists, skipping: $destinationPath"
    }
}

# Replace URL string for logos in the JSON with the relative path
if (-not (Test-Path -LiteralPath $OutputPath)) {
    throw "File not found: $OutputPath"
}

# Read the file as raw text so formatting is preserved as much as possible
$content = Get-Content -LiteralPath $OutputPath -Raw

# Replace only the companyLogoUrl prefix
$content = $content -replace `
    '(?m)^(\s*"companyLogoUrl"\s*:\s*")https://', `
    '${1}https://raw.githubusercontent.com/repasscloud/aethon-software-import-data/refs/heads/main/'

# Save back to the same file
Set-Content -LiteralPath $OutputPath -Value $content -Encoding utf8
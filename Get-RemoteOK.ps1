[CmdletBinding()]
param(
    [string]$ApiUrl = "https://remoteok.com/api",
    [string]$OutputPath = (Join-Path -Path $PSScriptRoot -ChildPath "remoteok-import.json")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

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

$CountryAliases = @{
    "UK"                       = "United Kingdom"
    "U.K."                     = "United Kingdom"
    "England"                  = "United Kingdom"
    "Scotland"                 = "United Kingdom"
    "Wales"                    = "United Kingdom"
    "USA"                      = "United States"
    "US"                       = "United States"
    "U.S."                     = "United States"
    "U.S.A."                   = "United States"
    "United States of America" = "United States"
    "UAE"                      = "United Arab Emirates"
    "KSA"                      = "Saudi Arabia"
    "Remote"                   = $null
    "Anywhere"                 = "Worldwide"
    "Global"                   = "Worldwide"
    "Worldwide"                = "Worldwide"
    "EMEA"                     = "EMEA"
    "APAC"                     = "APAC"
    "ANZ"                      = "Oceania"
    "LATAM"                    = "LatinAmerica"
    "EU"                       = "Europe"
}

$UsStateCodes = @(
    "AL","AK","AZ","AR","CA","CO","CT","DE","FL","GA","HI","ID","IL","IN","IA","KS","KY","LA",
    "ME","MD","MA","MI","MN","MS","MO","MT","NE","NV","NH","NJ","NM","NY","NC","ND","OH","OK",
    "OR","PA","RI","SC","SD","TN","TX","UT","VT","VA","WA","WV","WI","WY","DC"
)

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

function Repair-Encoding {
    <#
    .SYNOPSIS
        Fixes mojibake produced when UTF-8 bytes were stored/served as if they were
        Latin-1 characters (a common RemoteOK API issue).

    .DESCRIPTION
        Treats each character as a raw Latin-1 byte, then re-decodes the resulting
        byte array as strict UTF-8.  If the byte sequence is not valid UTF-8 the
        original string is returned unchanged, so pure-ASCII and genuinely Latin-1
        content is always safe.

        Examples of what this repairs:
          â€" (E2 80 94 as Latin-1 chars)  →  — (U+2014 em dash)
          â€™ (E2 80 99)                   →  ' (U+2019 right single quote)
          Â®  (C2 AE)                       →  ® (U+00AE registered sign)
    #>
    param(
        [AllowNull()]
        [string]$Text
    )

    if ([string]::IsNullOrEmpty($Text)) {
        return $Text
    }

    $bytes = [System.Text.Encoding]::Latin1.GetBytes($Text)
    try {
        $strict   = [System.Text.UTF8Encoding]::new($false, $true)  # throwOnInvalidBytes = true
        $repaired = $strict.GetString($bytes)
        return $repaired
    }
    catch {
        # Byte sequence is not valid UTF-8 — content is genuinely ASCII/Latin-1.
        return $Text
    }
}

function Normalize-String {
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

    return Repair-Encoding $text.Trim()
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

function Strip-HtmlToPlainText {
    param(
        [AllowNull()]
        [string]$Html
    )

    if ([string]::IsNullOrWhiteSpace($Html)) {
        return $null
    }

    $text = $Html
    $text = $text -replace '(?i)<\s*br\s*/?\s*>', "`n"
    $text = $text -replace '(?i)</\s*p\s*>', "`n"
    $text = $text -replace '(?i)</\s*div\s*>', "`n"
    $text = $text -replace '(?i)</\s*li\s*>', "`n"
    $text = $text -replace '(?i)<\s*li[^>]*>', ' - '
    $text = $text -replace '(?i)<\s*/?\s*ul\s*>', "`n"
    $text = $text -replace '(?i)<\s*/?\s*ol\s*>', "`n"

    $text = [regex]::Replace($text, '<[^>]+>', ' ')
    $text = [System.Net.WebUtility]::HtmlDecode($text)

    $text = $text -replace '\u00A0', ' '
    $text = $text -replace '\r', ''
    $text = $text -replace '[ \t]+', ' '
    $text = $text -replace '\n\s+', "`n"
    $text = $text -replace '\n{3,}', "`n`n"
    $text = $text -replace '[ ]{2,}', ' '

    return $text.Trim()
}

function Get-PlainTextSummary {
    param(
        [AllowNull()]
        [string]$Description,
        [int]$MaxLength = 100
    )

    $plain = Strip-HtmlToPlainText -Html $Description
    if ([string]::IsNullOrWhiteSpace($plain)) {
        return $null
    }

    # Summary is a single-line field — collapse newlines and excess whitespace to spaces.
    $plain = $plain -replace '[\r\n]+', ' ' -replace '\s{2,}', ' '
    $plain = $plain.Trim()

    if ($plain.Length -le $MaxLength) {
        return $plain
    }

    return $plain.Substring(0, $MaxLength)
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
        $number = [decimal]$text
        if ($number -le 0) {
            return $null
        }

        return $number
    }
    catch {
        return $null
    }
}

function Normalize-LocationText {
    param(
        [AllowNull()]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    $text = [System.Net.WebUtility]::HtmlDecode($Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    $text = $text -replace '\s+', ' '
    return $text.Trim()
}

function Get-LocationParts {
    param(
        [AllowNull()]
        [string]$LocationText
    )

    if ([string]::IsNullOrWhiteSpace($LocationText)) {
        return @()
    }

    @(
        $LocationText -split '[,|/;]' |
        ForEach-Object { $_.Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )
}

function Get-CountryFromLocation {
    param(
        [AllowNull()]
        [string]$LocationText
    )

    $normalizedLocation = Normalize-LocationText -Value $LocationText
    if ([string]::IsNullOrWhiteSpace($normalizedLocation)) {
        return $null
    }

    $parts = @(Get-LocationParts -LocationText $normalizedLocation)

    foreach ($part in $parts) {
        $candidate = Normalize-CountryName -Value $part
        if ($null -ne $candidate) {
            if ($candidate -eq "Worldwide" -or
                $candidate -eq "EMEA" -or
                $candidate -eq "APAC" -or
                $candidate -eq "Europe" -or
                $candidate -eq "LatinAmerica" -or
                $candidate -eq "NorthAmerica" -or
                $candidate -eq "Oceania") {
                return $null
            }

            if ($CountryToRegions.ContainsKey($candidate)) {
                return $candidate
            }
        }
    }

    $fullCandidate = Normalize-CountryName -Value $normalizedLocation
    if ($null -ne $fullCandidate -and $CountryToRegions.ContainsKey($fullCandidate)) {
        return $fullCandidate
    }

    if ($parts.Count -ge 2) {
        $last = $parts[-1].ToUpperInvariant()
        if ($UsStateCodes -contains $last) {
            return "United States"
        }
    }

    foreach ($country in $CountryToRegions.Keys) {
        if ($normalizedLocation -match "(?i)(^|[,\s\-/()])$([regex]::Escape($country))($|[,\s\-/()])") {
            return $country
        }
    }

    if ($normalizedLocation -match '(?i)\bUSA\b|\bU\.S\.A\.\b|\bUnited States\b|\bUS\b') {
        return "United States"
    }

    if ($normalizedLocation -match '(?i)\bUK\b|\bU\.K\.\b|\bUnited Kingdom\b|\bEngland\b|\bScotland\b|\bWales\b') {
        return "United Kingdom"
    }

    if ($normalizedLocation -match '(?i)\bUAE\b|\bUnited Arab Emirates\b') {
        return "United Arab Emirates"
    }

    return $null
}

function Get-RegionsFromLocationText {
    param(
        [AllowNull()]
        [string]$LocationText,
        [AllowNull()]
        [string]$Country
    )

    $regions = New-Object System.Collections.Generic.List[string]
    $normalizedLocation = Normalize-LocationText -Value $LocationText

    if (-not [string]::IsNullOrWhiteSpace($normalizedLocation)) {
        if ($normalizedLocation -match '(?i)\bworldwide\b|\bglobal\b|\banywhere\b') {
            if (-not $regions.Contains("Worldwide")) { [void]$regions.Add("Worldwide") }
        }

        if ($normalizedLocation -match '(?i)\bEMEA\b') {
            foreach ($r in @("Europe", "MiddleEast", "Africa")) {
                if (-not $regions.Contains($r)) { [void]$regions.Add($r) }
            }
        }

        if ($normalizedLocation -match '(?i)\bAPAC\b') {
            foreach ($r in @("Asia", "Oceania")) {
                if (-not $regions.Contains($r)) { [void]$regions.Add($r) }
            }
        }

        if ($normalizedLocation -match '(?i)\bLATAM\b|\bLatin America\b') {
            if (-not $regions.Contains("LatinAmerica")) { [void]$regions.Add("LatinAmerica") }
        }

        if ($normalizedLocation -match '(?i)\bNorth America\b') {
            if (-not $regions.Contains("NorthAmerica")) { [void]$regions.Add("NorthAmerica") }
        }

        if ($normalizedLocation -match '(?i)\bEurope\b|\bEU\b') {
            if (-not $regions.Contains("Europe")) { [void]$regions.Add("Europe") }
        }

        if ($normalizedLocation -match '(?i)\bAsia\b') {
            if (-not $regions.Contains("Asia")) { [void]$regions.Add("Asia") }
        }

        if ($normalizedLocation -match '(?i)\bAfrica\b') {
            if (-not $regions.Contains("Africa")) { [void]$regions.Add("Africa") }
        }

        if ($normalizedLocation -match '(?i)\bMiddle East\b') {
            if (-not $regions.Contains("MiddleEast")) { [void]$regions.Add("MiddleEast") }
        }

        if ($normalizedLocation -match '(?i)\bOceania\b|\bANZ\b|\bAustralia\b|\bNew Zealand\b') {
            if (-not $regions.Contains("Oceania")) { [void]$regions.Add("Oceania") }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($Country) -and $CountryToRegions.ContainsKey($Country)) {
        foreach ($region in $CountryToRegions[$Country]) {
            if (-not $regions.Contains($region)) {
                [void]$regions.Add($region)
            }
        }
    }

    return @($regions)
}

function Get-EmploymentType {
    param(
        [AllowNull()]
        [string[]]$Tags,
        [AllowNull()]
        [string]$Title,
        [AllowNull()]
        [string]$DescriptionPlain
    )

    $text = ((@($Tags) + @($Title) + @($DescriptionPlain)) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ' '

    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    switch -Regex ($text) {
        '(?i)\bfull[ -]?time\b' { return "FullTime" }
        '(?i)\bpart[ -]?time\b' { return "PartTime" }
        '(?i)\bcontract\b|\bcontractor\b|\bfreelance\b' { return "Contract" }
        '(?i)\btemporary\b|\btemp\b' { return "Temporary" }
        '(?i)\bintern\b|\binternship\b' { return "Internship" }
        '(?i)\bcasual\b' { return "Casual" }
        default { return $null }
    }
}

function Get-Category {
    param(
        [AllowNull()]
        [string[]]$Tags,
        [AllowNull()]
        [string]$Title,
        [AllowNull()]
        [string]$DescriptionPlain
    )

    $text = ((@($Tags) + @($Title) + @($DescriptionPlain)) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ' '

    if ([string]::IsNullOrWhiteSpace($text)) {
        return "Other"
    }

    switch -Regex ($text) {
        '(?i)\bsoftware\b|\bdeveloper\b|\bdevelopment\b|\bprogramming\b|\bengineer\b|\bengineering\b|\bqa\b|\bdevops\b|\bsre\b|\bcloud\b|\bsecurity\b|\bdata engineer\b|\bbackend\b|\bfront[ -]?end\b|\bfull[ -]?stack\b|\bmobile\b|\bplatform\b|\binfrastructure\b|\btechnical\b' { return "ITSoftware" }
        '(?i)\bdata scientist\b|\bmachine learning\b|\bml\b|\bai\b|\banalytics\b|\bbi\b|\bdata analyst\b' { return "ITSoftware" }
        '(?i)\bdesign\b|\bux\b|\bui\b|\bgraphic\b|\bproduct design\b|\bvisual design\b' { return "Design" }
        '(?i)\bmarketing\b|\bseo\b|\bcontent\b|\bbrand\b|\bgrowth\b|\bdemand gen\b|\bdigital marketing\b|\bsocial media\b|\bcopywriter\b' { return "Marketing" }
        '(?i)\bsales\b|\baccount executive\b|\bbusiness development\b|\bsdr\b|\bbdr\b|\bpartnerships\b' { return "Sales" }
        '(?i)\bcustomer success\b|\bcustomer service\b|\bsupport\b|\btechnical support\b|\bhelp desk\b' { return "CustomerService" }
        '(?i)\bhuman resources\b|\bhr\b|\btalent\b|\bpeople ops\b|\bpeople operations\b' { return "HumanResources" }
        '(?i)\bfinance\b|\baccounting\b|\bbookkeep\b|\bcontroller\b|\bfinancial analyst\b' { return "FinanceInsurance" }
        '(?i)\bbanking\b' { return "Banking" }
        '(?i)\beducation\b|\bteaching\b|\bteacher\b|\btraining\b|\binstructional\b' { return "Education" }
        '(?i)\bhealthcare\b|\bmedical\b|\bnursing\b|\bclinical\b|\btherapist\b|\bpsychologist\b' { return "Healthcare" }
        '(?i)\blegal\b|\bparalegal\b|\bcounsel\b|\battorney\b' { return "LegalServices" }
        '(?i)\brecruitment\b|\brecruiter\b|\bstaffing\b|\bsourcer\b' { return "Recruitment" }
        '(?i)\bresearch\b' { return "Research" }
        '(?i)\bscience\b|\bscientist\b' { return "Science" }
        '(?i)\bmedia\b|\bjournalism\b|\beditorial\b|\bpublishing\b|\bwriter\b' { return "MediaJournalism" }
        '(?i)\bpublic relations\b|\bpr\b' { return "PublicRelations" }
        '(?i)\badvertising\b' { return "AdvertisingPR" }
        '(?i)\badministration\b|\badmin\b|\bsecretarial\b|\boffice manager\b|\bexecutive assistant\b' { return "AdminSecretarial" }
        '(?i)\bhospitality\b|\bhotel\b|\brestaurant\b' { return "Hospitality" }
        '(?i)\bcatering\b' { return "Catering" }
        '(?i)\bretail\b|\becommerce\b|\be-commerce\b|\bmerchant\b' { return "Retail" }
        '(?i)\blogistics\b|\bsupply chain\b|\bwarehouse\b|\boperations\b' { return "Logistics" }
        '(?i)\btransport\b|\bdistribution\b|\bdriver\b' { return "TransportDistribution" }
        '(?i)\bmanufacturing\b|\bproduction\b' { return "Manufacturing" }
        '(?i)\bconstruction\b|\bbuilding\b' { return "BuildingConstruction" }
        '(?i)\bproperty\b|\breal estate\b' { return "PropertyRealEstate" }
        '(?i)\bpharma\b|\bpharmaceutical\b' { return "Pharmaceuticals" }
        '(?i)\btelecommunications\b|\btelecom\b|\bisp\b' { return "TelecommunicationsISP" }
        '(?i)\bgovernment\b|\bpublic sector\b' { return "Government" }
        '(?i)\bexecutive\b|\bleadership\b|\bmanagement\b|\bdirector\b|\bvp\b|\bchief\b|\bhead of\b' { return "ExecutiveManagement" }
        '(?i)\bgraduate\b|\bentry level\b|\bjunior\b' { return "GraduateRoles" }
        '(?i)\btourism\b|\btravel\b' { return "Tourism" }
        '(?i)\butilities\b|\benergy\b' { return "UtilitiesEnergy" }
        '(?i)\bagriculture\b|\bfishing\b|\bforestry\b' { return "AgricultureFishingForestry" }
        '(?i)\bmining\b|\bresources\b' { return "MiningResources" }
        '(?i)\bsport\b|\brecreation\b|\bfitness\b' { return "SportRecreation" }
        '(?i)\barts\b|\bcreative\b' { return "Arts" }
        '(?i)\bcharity\b|\bnonprofit\b|\bnon-profit\b|\bngo\b' { return "Charity" }
        '(?i)\bfood\b|\bbeverage\b' { return "FoodBeverage" }
        '(?i)\bautomotive\b|\bautomobile\b' { return "Automobile" }
        '(?i)\baerospace\b|\baviation\b' { return "Aerospace" }
        '(?i)\bveterinary\b' { return "Veterinary" }
        '(?i)\bsocial work\b|\bcommunity\b' { return "SocialWork" }
        default { return "Other" }
    }
}

function Get-CompanyLogoUrl {
    <#
    .SYNOPSIS
        Scrapes the RemoteOK job page for the company logo URL.

    .DESCRIPTION
        The RemoteOK API does not expose a logo field, but each job page contains
        an <img> element like:

            data-src="https://resizeapi.com/resize-cgi/image/format=auto,...,quality=80/
                      https://r2.remoteok.com/jobs/<hash>.png" class="logo"

        This function fetches the page, extracts the actual logo URL that follows the
        resize-API prefix, and returns it.  Returns $null on any error or if the
        element is not found.
    #>
    param(
        [AllowNull()]
        [string]$JobPageUrl
    )

    if ([string]::IsNullOrWhiteSpace($JobPageUrl)) { return $null }

    try {
        $response = Invoke-WebRequest -Uri $JobPageUrl -UseBasicParsing -TimeoutSec 15 `
            -Headers @{ 'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36' }

        $html = [System.Text.Encoding]::UTF8.GetString($response.RawContentStream.ToArray())

        # Match the actual logo URL embedded inside the resize-API wrapper.
        # Pattern: data-src="https://resizeapi.com/.../{LOGO_URL}" ... class="logo"
        $m = [regex]::Match(
            $html,
            'data-src="https?://resizeapi\.com/[^"]*/(https?://[^"]+)"[^>]*class="logo"'
        )
        if ($m.Success) {
            return $m.Groups[1].Value.Trim()
        }
    }
    catch {
        # Logo is optional — ignore all errors silently
    }

    return $null
}

function Get-SafeUrl {
    param(
        [Parameter(Mandatory)]
        [object]$Job
    )

    $applyUrl = Normalize-String (Get-ObjectPropertyValue -InputObject $Job -PropertyName "apply_url")
    if (-not [string]::IsNullOrWhiteSpace($applyUrl)) {
        return $applyUrl
    }

    return Normalize-String (Get-ObjectPropertyValue -InputObject $Job -PropertyName "url")
}

Write-Host "Fetching Remote OK jobs from $ApiUrl"

try {
    # Invoke-RestMethod can fall back to ISO-8859-1 when the server omits a charset
    # in its Content-Type header, which turns multi-byte UTF-8 characters (em dashes,
    # curly quotes, ® etc.) into mojibake.  Read the raw bytes and force UTF-8 decoding.
    $webResponse  = Invoke-WebRequest -Uri $ApiUrl -Method Get -UseBasicParsing `
        -Headers @{ 'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36' }
    $responseText = [System.Text.Encoding]::UTF8.GetString($webResponse.RawContentStream.ToArray())
    $response     = $responseText | ConvertFrom-Json
}
catch {
    throw "Failed to fetch jobs from '$ApiUrl': $_"
}

if ($null -eq $response) {
    throw "The API returned no data."
}

$jobs = @(
    $response |
    Where-Object {
        -not [string]::IsNullOrWhiteSpace((Get-ObjectPropertyValue -InputObject $_ -PropertyName "id")) -and
        -not [string]::IsNullOrWhiteSpace((Get-ObjectPropertyValue -InputObject $_ -PropertyName "position")) -and
        -not [string]::IsNullOrWhiteSpace((Get-ObjectPropertyValue -InputObject $_ -PropertyName "company"))
    }
)

if ($jobs.Count -eq 0) {
    throw "The API response did not contain any job records."
}

$transformed = foreach ($job in $jobs) {
    $descriptionHtml  = Normalize-String (Get-ObjectPropertyValue -InputObject $job -PropertyName "description")
    $descriptionPlain = Strip-HtmlToPlainText -Html $descriptionHtml
    $locationText     = Normalize-LocationText -Value (Get-ObjectPropertyValue -InputObject $job -PropertyName "location")
    $locationCountry  = Get-CountryFromLocation -LocationText $locationText
    $regions          = @(Get-RegionsFromLocationText -LocationText $locationText -Country $locationCountry)

    $tags = @(
        Normalize-ToArray -Value (Get-ObjectPropertyValue -InputObject $job -PropertyName "tags") |
        ForEach-Object { Normalize-String $_ } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )

    $title = Normalize-String (Get-ObjectPropertyValue -InputObject $job -PropertyName "position")

    $countriesList = New-Object System.Collections.Generic.List[string]
    if (-not [string]::IsNullOrWhiteSpace($locationCountry)) { $countriesList.Add($locationCountry) }

    # The API never populates company_logo, so scrape the logo from the job page.
    # Always use the remoteok.com `url` field — apply_url may point to an external site.
    $jobPageUrl    = Normalize-String (Get-ObjectPropertyValue -InputObject $job -PropertyName "url")
    $companyLogo   = Get-CompanyLogoUrl -JobPageUrl $jobPageUrl
    Write-Host "  Logo for '$( Normalize-String (Get-ObjectPropertyValue -InputObject $job -PropertyName "company") )': $( if ($companyLogo) { $companyLogo } else { '(none)' } )"

    $jobObj = [PSCustomObject]@{
        sourceSite             = "remoteok.com"
        externalId             = Normalize-String (Get-ObjectPropertyValue -InputObject $job -PropertyName "id")
        companyName            = Normalize-String (Get-ObjectPropertyValue -InputObject $job -PropertyName "company")
        companyLogoUrl         = $companyLogo

        title                  = $title
        description            = $descriptionHtml
        workplaceType          = "Remote"
        employmentType         = Get-EmploymentType -Tags $tags -Title $title -DescriptionPlain $descriptionPlain

        externalApplicationUrl = Get-SafeUrl -Job $job

        category               = Get-Category -Tags $tags -Title $title -DescriptionPlain $descriptionPlain
        keywords               = if ($tags.Count -gt 0) { $tags -join ',' } else { $null }
        regions                = @($regions)
        countries              = $countriesList

        summary                = Get-PlainTextSummary -Description $descriptionHtml -MaxLength 100
        requirements           = $null
        benefits               = $null
        department             = $null

        salaryFrom             = Convert-SafeDecimal -Value (Get-ObjectPropertyValue -InputObject $job -PropertyName "salary_min")
        salaryTo               = Convert-SafeDecimal -Value (Get-ObjectPropertyValue -InputObject $job -PropertyName "salary_max")
        salaryCurrency         = $null

        publishedUtc           = Convert-ToUtcIsoString -Value (Get-ObjectPropertyValue -InputObject $job -PropertyName "date")
        postingExpiresUtc      = $null

        locationText           = $locationText
        locationCity           = $null
        locationState          = $null
        locationCountry        = $locationCountry
        locationCountryCode    = $null
        locationLatitude       = $null
        locationLongitude      = $null

        slug                   = Normalize-String (Get-ObjectPropertyValue -InputObject $job -PropertyName "slug")
    }

    # employmentType is a non-nullable enum in the import API — omit the property
    # entirely when unknown so the API uses its default rather than rejecting with 400.
    if ($null -eq $jobObj.employmentType) {
        $jobObj.PSObject.Properties.Remove('employmentType')
    }

    $jobObj
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

# ---------------------------------------------------------------------------
# Phase 2: Download company logos and rewrite companyLogoUrl to GitHub raw URLs
# ---------------------------------------------------------------------------

$logoLocalRoot  = Join-Path -Path $PSScriptRoot -ChildPath "remoteok.com"
$githubRawBase  = "https://raw.githubusercontent.com/repasscloud/aethon-software-import-data/refs/heads/main/remoteok.com"

if (-not (Test-Path $logoLocalRoot)) {
    New-Item -ItemType Directory -Path $logoLocalRoot -Force | Out-Null
}

Write-Host ""
Write-Host "Downloading company logos..."

# Re-read the JSON we just wrote so we work with the serialized form
$importData = Get-Content -Path $OutputPath -Raw -Encoding UTF8 | ConvertFrom-Json

$anyUpdated = $false

foreach ($item in $importData) {
    $logoUrl = $item.companyLogoUrl
    if ([string]::IsNullOrWhiteSpace($logoUrl)) { continue }

    # Derive the path component (e.g. /jobs/abc.jpg) from whatever URL was scraped
    try { $uri = [System.Uri]$logoUrl } catch { continue }
    $relativePath = $uri.AbsolutePath          # always starts with /
    $relativePathClean = $relativePath.TrimStart('/')   # jobs/abc.jpg

    # Build local file path
    $localPath = Join-Path -Path $logoLocalRoot -ChildPath $relativePathClean

    # Ensure sub-directory exists
    $localDir = Split-Path -Path $localPath -Parent
    if (-not (Test-Path $localDir)) {
        New-Item -ItemType Directory -Path $localDir -Force | Out-Null
    }

    # Download only if not already cached
    if (-not (Test-Path $localPath)) {
        try {
            Invoke-WebRequest -Uri $logoUrl -OutFile $localPath -UseBasicParsing -TimeoutSec 30 `
                -Headers @{ 'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36' }
            Write-Host "  Downloaded : $relativePathClean"
        }
        catch {
            Write-Warning "  Failed to download logo '$logoUrl': $_"
            continue   # leave the original URL intact if download failed
        }
    }
    else {
        Write-Host "  Cached      : $relativePathClean"
    }

    # Rewrite the URL in the in-memory object
    $item.companyLogoUrl = "$githubRawBase/$relativePathClean"
    $anyUpdated = $true
}

if ($anyUpdated) {
    $updatedJson = @($importData) | ConvertTo-Json -Depth 20
    try {
        $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
        [System.IO.File]::WriteAllText($OutputPath, $updatedJson, $utf8NoBom)
        Write-Host "Updated companyLogoUrl entries rewritten to GitHub raw URLs."
    }
    catch {
        throw "Failed to write updated JSON to '$OutputPath': $_"
    }
}
else {
    Write-Host "No logo URLs required updating."
}
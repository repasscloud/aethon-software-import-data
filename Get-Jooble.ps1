[CmdletBinding()]
param(
    [string]$OutputDirectory = $PSScriptRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$JOOBLE_API_KEY = $env:JOOBLE_API_KEY
$JOOBLE_API_URL = "https://jooble.org/api/"

if ([string]::IsNullOrWhiteSpace($JOOBLE_API_KEY)) {
    Write-Warning "JOOBLE_API_KEY is missing — writing empty output and skipping."
    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    [System.IO.File]::WriteAllText($OutputPath, "[]", $utf8NoBom)
    exit 0
}

# ---------------------------------------------------------------------------
# Region / country lookup tables (shared with other import scripts)
# ---------------------------------------------------------------------------

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
    NorthAmerica = @("Canada","Mexico","United States")
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

$CountryList = @(
    $RegionCountryMap.Values |
    ForEach-Object { $_ } |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
    Sort-Object -Unique
)

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

# ---------------------------------------------------------------------------
# Helper functions
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

function Strip-HtmlToPlainText {
    param([AllowNull()][string]$Html)
    if ([string]::IsNullOrWhiteSpace($Html)) { return $null }
    $t = $Html
    $t = $t -replace '(?i)<\s*br\s*/?\s*>',  "`n"
    $t = $t -replace '(?i)</\s*p\s*>',         "`n"
    $t = $t -replace '(?i)</\s*div\s*>',       "`n"
    $t = $t -replace '(?i)</\s*li\s*>',        "`n"
    $t = $t -replace '(?i)<\s*li[^>]*>',       ' - '
    $t = [regex]::Replace($t, '<[^>]+>', ' ')
    $t = [System.Net.WebUtility]::HtmlDecode($t)
    $t = $t -replace '\u00A0', ' '
    $t = $t -replace '[ \t]+', ' '
    $t = $t -replace '\n\s+', "`n"
    $t = $t -replace '\n{3,}', "`n`n"
    return $t.Trim()
}

function Get-PlainTextSummary {
    param([AllowNull()][string]$Text, [int]$MaxLength = 100)
    $plain = Strip-HtmlToPlainText -Html $Text
    if ([string]::IsNullOrWhiteSpace($plain)) { return $null }
    $plain = ($plain -replace '[\r\n]+', ' ' -replace '\s{2,}', ' ').Trim()
    if ($plain.Length -le $MaxLength) { return $plain }
    return $plain.Substring(0, $MaxLength)
}

function Normalize-CountryName {
    param([AllowNull()][string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $t = $Value.Trim()
    if ($CountryAliases.ContainsKey($t)) { return $CountryAliases[$t] }
    return $t
}

function Get-CountryFromLocation {
    param([AllowNull()][string]$LocationText)
    if ([string]::IsNullOrWhiteSpace($LocationText)) { return $null }
    $loc = $LocationText.Trim()

    # Split on common separators and test each part
    $parts = @($loc -split '[,|/;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    foreach ($part in $parts) {
        $c = Normalize-CountryName -Value $part
        if ($null -ne $c -and $CountryToRegions.ContainsKey($c)) { return $c }
    }

    # Test the whole string
    $full = Normalize-CountryName -Value $loc
    if ($null -ne $full -and $CountryToRegions.ContainsKey($full)) { return $full }

    # US state code as last part
    if ($parts.Count -ge 2) {
        $last = $parts[-1].ToUpperInvariant()
        if ($UsStateCodes -contains $last) { return "United States" }
    }

    # Substring scan
    foreach ($country in $CountryToRegions.Keys) {
        if ($loc -match "(?i)(^|[,\s\-/()])$([regex]::Escape($country))($|[,\s\-/()])") {
            return $country
        }
    }

    if ($loc -match '(?i)\bUSA\b|\bU\.S\.A\.\b|\bUnited States\b') { return "United States" }
    if ($loc -match '(?i)\bUK\b|\bU\.K\.\b|\bUnited Kingdom\b|\bEngland\b') { return "United Kingdom" }
    if ($loc -match '(?i)\bAustralia\b|\bAUS\b|\bNSW\b|\bVIC\b|\bQLD\b|\bSA\b|\bWA\b|\bTAS\b|\bACT\b|\bNT\b') { return "Australia" }
    return $null
}

function Get-RegionsFromLocationText {
    param([AllowNull()][string]$LocationText, [AllowNull()][string]$Country)
    $regions = New-Object System.Collections.Generic.List[string]
    if (-not [string]::IsNullOrWhiteSpace($LocationText)) {
        $loc = $LocationText
        if ($loc -match '(?i)\bworldwide\b|\bglobal\b|\banywhere\b|\bremote\b') {
            # remote alone doesn't imply worldwide — only if truly global language
            if ($loc -match '(?i)\bworldwide\b|\bglobal\b|\banywhere\b') {
                if (-not $regions.Contains("Worldwide")) { [void]$regions.Add("Worldwide") }
            }
        }
        if ($loc -match '(?i)\bEMEA\b') {
            foreach ($r in @("Europe","MiddleEast","Africa")) {
                if (-not $regions.Contains($r)) { [void]$regions.Add($r) }
            }
        }
        if ($loc -match '(?i)\bAPAC\b') {
            foreach ($r in @("Asia","Oceania")) {
                if (-not $regions.Contains($r)) { [void]$regions.Add($r) }
            }
        }
        if ($loc -match '(?i)\bLATAM\b|\bLatin America\b') {
            if (-not $regions.Contains("LatinAmerica")) { [void]$regions.Add("LatinAmerica") }
        }
        if ($loc -match '(?i)\bNorth America\b') {
            if (-not $regions.Contains("NorthAmerica")) { [void]$regions.Add("NorthAmerica") }
        }
        if ($loc -match '(?i)\bEurope\b|\bEU\b') {
            if (-not $regions.Contains("Europe")) { [void]$regions.Add("Europe") }
        }
        if ($loc -match '(?i)\bAsia\b') {
            if (-not $regions.Contains("Asia")) { [void]$regions.Add("Asia") }
        }
        if ($loc -match '(?i)\bAfrica\b') {
            if (-not $regions.Contains("Africa")) { [void]$regions.Add("Africa") }
        }
        if ($loc -match '(?i)\bMiddle East\b') {
            if (-not $regions.Contains("MiddleEast")) { [void]$regions.Add("MiddleEast") }
        }
        if ($loc -match '(?i)\bOceania\b|\bANZ\b|\bAustralia\b|\bNew Zealand\b|\bNSW\b|\bVIC\b|\bQLD\b|\bSouth Australia\b|\bSA\b|\bWA\b|\bTAS\b|\bACT\b|\bNT\b') {
            if (-not $regions.Contains("Oceania")) { [void]$regions.Add("Oceania") }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($Country) -and $CountryToRegions.ContainsKey($Country)) {
        foreach ($r in $CountryToRegions[$Country]) {
            if (-not $regions.Contains($r)) { [void]$regions.Add($r) }
        }
    }

    # Default to Oceania for South Australia searches with no region match
    if ($regions.Count -eq 0) {
        [void]$regions.Add("Oceania")
    }

    return @($regions)
}

# ---------------------------------------------------------------------------
# Employment type — maps Jooble "type" field to EmploymentType enum values
# ---------------------------------------------------------------------------

function Convert-EmploymentType {
    param([AllowNull()][string]$Type, [AllowNull()][string]$Title)

    $combined = ((@($Type) + @($Title)) | Where-Object { $_ }) -join ' '
    if ([string]::IsNullOrWhiteSpace($combined)) { return $null }

    switch -Regex ($combined) {
        '(?i)\bfull[\s\-]?time\b'               { return "FullTime"   }
        '(?i)\bpart[\s\-]?time\b'               { return "PartTime"   }
        '(?i)\bcontract\b|\bcontractor\b|\bfreelance\b' { return "Contract"   }
        '(?i)\btemporary\b|\btemp\b'            { return "Temporary"  }
        '(?i)\bintern\b|\binternship\b'         { return "Internship" }
        '(?i)\bcasual\b'                        { return "Casual"     }
    }
    return $null
}

# ---------------------------------------------------------------------------
# WorkplaceType
# ---------------------------------------------------------------------------

function Get-WorkplaceType {
    param([AllowNull()][string]$LocationText, [AllowNull()][string]$Type)
    $combined = ((@($LocationText) + @($Type)) | Where-Object { $_ }) -join ' '
    if ($combined -match '(?i)\bremote\b|\bwork from home\b|\bwfh\b') { return "Remote" }
    if ($combined -match '(?i)\bhybrid\b')                             { return "Hybrid"  }
    return "OnSite"
}

# ---------------------------------------------------------------------------
# Category — inferred from title + snippet
# ---------------------------------------------------------------------------

function Get-JobCategory {
    param([AllowNull()][string]$Title, [AllowNull()][string]$SnippetPlain)

    $text = ((@($Title) + @($SnippetPlain)) | Where-Object { $_ }) -join ' '
    if ([string]::IsNullOrWhiteSpace($text)) { return "Other" }

    switch -Regex ($text) {
        '(?i)\bsoftware\b|\bdeveloper\b|\bdevelopment\b|\bprogramming\b|\bengineer(ing)?\b|\bqa\b|\bdevops\b|\bsre\b|\bcloud\b|\bbackend\b|\bfront[\s\-]?end\b|\bfull[\s\-]?stack\b|\bmobile\b|\bplatform\b|\binfrastructure\b|\btechnical\b|\bcyber\b|\bsecurity\b' { return "ITSoftware" }
        '(?i)\bdata scientist\b|\bmachine learning\b|\bml\b|\b\bai\b|\banalytics\b|\bdata analyst\b' { return "ITSoftware" }
        '(?i)\bdesign\b|\bux\b|\bui\b|\bgraphic\b|\bvisual design\b'                          { return "Design" }
        '(?i)\bmarketing\b|\bseo\b|\bcontent\b|\bbrand\b|\bgrowth\b|\bdigital marketing\b|\bsocial media\b|\bcopywriter\b' { return "Marketing" }
        '(?i)\bsales\b|\baccount executive\b|\bbusiness development\b|\bsdr\b|\bbdr\b|\bpartnerships\b' { return "Sales" }
        '(?i)\bcustomer success\b|\bcustomer service\b|\bsupport\b|\bhelp desk\b'             { return "CustomerService" }
        '(?i)\bhuman resources\b|\bhr\b|\btalent\b|\bpeople ops\b'                            { return "HumanResources" }
        '(?i)\bfinance\b|\baccounting\b|\bbookkeep\b|\bcontroller\b|\bfinancial analyst\b'    { return "Accounting" }
        '(?i)\bbanking\b'                                                                       { return "Banking" }
        '(?i)\beducation\b|\bteaching\b|\bteacher\b|\btraining\b|\binstructional\b'           { return "Education" }
        '(?i)\bhealthcare\b|\bmedical\b|\bnursing\b|\bclinical\b|\btherapist\b'               { return "Healthcare" }
        '(?i)\blegal\b|\bparalegal\b|\bcounsel\b|\battorney\b'                                { return "LegalServices" }
        '(?i)\brecruitment\b|\brecruiter\b|\bstaffing\b|\bsourcer\b'                          { return "Recruitment" }
        '(?i)\bresearch\b'                                                                      { return "Research" }
        '(?i)\bscience\b|\bscientist\b'                                                        { return "Science" }
        '(?i)\bmedia\b|\bjournalism\b|\beditorial\b|\bpublishing\b|\bwriter\b'                { return "MediaJournalism" }
        '(?i)\bpublic relations\b|\bpr manager\b'                                              { return "PublicRelations" }
        '(?i)\badvertising\b'                                                                   { return "AdvertisingPR" }
        '(?i)\badministrat\b|\badmin\b|\bsecretarial\b|\boffice manager\b|\bexecutive assistant\b' { return "AdminSecretarial" }
        '(?i)\bhospitality\b|\bhotel\b|\brestaurant\b'                                        { return "Hospitality" }
        '(?i)\bretail\b|\becommerce\b|\be-commerce\b'                                         { return "Retail" }
        '(?i)\blogistics\b|\bsupply chain\b|\bwarehouse\b'                                    { return "Logistics" }
        '(?i)\btransport\b|\bdistribution\b|\bdriver\b'                                       { return "TransportDistribution" }
        '(?i)\bmanufacturing\b|\bproduction\b'                                                 { return "Manufacturing" }
        '(?i)\bconstruction\b|\bbuilding\b'                                                    { return "BuildingConstruction" }
        '(?i)\bproperty\b|\breal estate\b'                                                     { return "PropertyRealEstate" }
        '(?i)\bpharma\b|\bpharmaceutical\b'                                                    { return "Pharmaceuticals" }
        '(?i)\btelecommunications\b|\btelecom\b|\bisp\b'                                      { return "TelecommunicationsISP" }
        '(?i)\bgovernment\b|\bpublic sector\b|\bpublic service\b'                             { return "Government" }
        '(?i)\bexecutive\b|\bdirector\b|\bvp\b|\bchief\b|\bhead of\b'                        { return "ExecutiveManagement" }
        '(?i)\bgraduate\b|\bentry level\b|\bjunior\b'                                         { return "GraduateRoles" }
        '(?i)\btourism\b|\btravel\b'                                                           { return "Tourism" }
        '(?i)\butilities\b|\benergy\b'                                                         { return "UtilitiesEnergy" }
        '(?i)\bagriculture\b|\bfishing\b|\bforestry\b'                                        { return "AgricultureFishingForestry" }
        '(?i)\bmining\b|\bresources\b'                                                         { return "MiningResources" }
        '(?i)\bsport\b|\brecreation\b|\bfitness\b'                                            { return "SportRecreation" }
        '(?i)\barts\b|\bcreative\b'                                                            { return "Arts" }
        '(?i)\bcharity\b|\bnonprofit\b|\bnon-profit\b|\bngo\b'                               { return "Charity" }
        '(?i)\bfood\b|\bbeverage\b'                                                            { return "FoodBeverage" }
        '(?i)\bautomotive\b|\bautomobile\b'                                                    { return "Automobile" }
        '(?i)\baerospace\b|\baviation\b'                                                       { return "Aerospace" }
        '(?i)\bveterinary\b'                                                                   { return "Veterinary" }
        '(?i)\bsocial work\b|\bcommunity\b'                                                   { return "SocialWork" }
        default                                                                                 { return "Other" }
    }
}

# ---------------------------------------------------------------------------
# Salary parsing — handles strings like "$50,000", "$50k – $70k", "80000"
# ---------------------------------------------------------------------------

function Parse-SalaryRange {
    param([AllowNull()][string]$SalaryText)

    if ([string]::IsNullOrWhiteSpace($SalaryText)) {
        return @{ From = $null; To = $null; Currency = $null }
    }

    # Detect currency symbol
    $currency = $null
    if ($SalaryText -match '\$') { $currency = "USD" }
    elseif ($SalaryText -match '£') { $currency = "GBP" }
    elseif ($SalaryText -match '€') { $currency = "EUR" }
    elseif ($SalaryText -match 'A\$') { $currency = "AUD" }

    # Skip non-annual rates (per hour, per day, per week)
    if ($SalaryText -match '(?i)\bper\s+(hour|hr|day|week)\b|\b/h\b|\bph\b') {
        return @{ From = $null; To = $null; Currency = $null }
    }

    # Normalise: remove currency symbols, commas, spaces; expand k → 000
    $clean = $SalaryText -replace '[£€$A]', '' -replace ',', '' -replace '\s', ''
    $clean = $clean -replace '(?i)(\d+)k', '$1000'

    # Try to extract one or two numbers (always force array so .Count is safe in strict mode)
    $nums = @([regex]::Matches($clean, '\d+(\.\d+)?') |
              ForEach-Object { [decimal]$_.Value } |
              Where-Object { $_ -ge 1000 })   # ignore tiny values (e.g. "2 years exp")

    if ($nums.Count -eq 0) {
        return @{ From = $null; To = $null; Currency = $null }
    }

    $from = $nums[0]
    $to   = if ($nums.Count -ge 2) { $nums[-1] } else { $null }

    # Sanity: if "to" is less than "from", discard it
    if ($null -ne $to -and $to -lt $from) { $to = $null }

    return @{ From = $from; To = $to; Currency = $currency }
}

# ---------------------------------------------------------------------------
# Slug
# ---------------------------------------------------------------------------

function ConvertTo-JobSlug {
    param([string]$Title, [string]$Id)
    $slugTitle = $Title.ToLower() `
        -replace '[^a-z0-9\s-]', '' `
        -replace '\s+', '-' `
        -replace '-+', '-'
    # Extract a short stable suffix from the id (last path segment or hash)
    $suffix = $Id -replace '^.*/', '' -replace '[^a-z0-9-]', '-' -replace '-+', '-' -replace '^-|-$', ''
    if ($suffix.Length -gt 20) { $suffix = $suffix.Substring(0, 20).TrimEnd('-') }
    $slug = "$slugTitle-$suffix"
    if ($slug.Length -gt 120) { $slug = $slug.Substring(0, 120).TrimEnd('-') }
    return $slug.ToLower()
}

# ---------------------------------------------------------------------------
# Description — wrap snippet HTML; ensure it fits the 20 000-char DB limit
# ---------------------------------------------------------------------------

function Build-Description {
    param([AllowNull()][string]$Snippet)
    if ([string]::IsNullOrWhiteSpace($Snippet)) {
        return "<p>Please visit the application URL for full job details.</p>"
    }

    # Snippet is already HTML from Jooble — strip any embedded newlines before wrapping
    $html = $Snippet.Trim() -replace '[\r\n]+', ' '
    if ($html -notmatch '(?i)^<(p|div|ul|ol|h[1-6])') {
        $html = "<p>$html</p>"
    }

    $maxLen  = 19800
    $trailer = "<p><em>Description truncated. View full details on the source site.</em></p>"
    if ($html.Length -gt $maxLen) {
        $cutAt = $maxLen
        while ($cutAt -gt 0 -and $html[$cutAt] -ne '>') { $cutAt-- }
        $html = $html.Substring(0, $cutAt + 1) + $trailer
    }

    return $html
}

# ---------------------------------------------------------------------------
# Fetch / transform / output per country
# ---------------------------------------------------------------------------

if (-not (Test-Path $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
}

$utf8NoBom = [System.Text.UTF8Encoding]::new($false)

foreach ($searchCountry in $CountryList) {

    $countryFilePart = $searchCountry.ToLowerInvariant() `
        -replace '[^a-z0-9]+', '-' `
        -replace '^-|-$', ''

    $currentOutputPath = Join-Path -Path $OutputDirectory -ChildPath ("jooble_{0}-import.json" -f $countryFilePart)

    Write-Host "Fetching Jooble jobs for: $searchCountry"

    $requestBody  = @{ location = $searchCountry } | ConvertTo-Json -Compress
    $requestBytes = [System.Text.Encoding]::UTF8.GetBytes($requestBody)

    $request = [System.Net.HttpWebRequest]::Create($JOOBLE_API_URL + $JOOBLE_API_KEY)
    $request.Method        = "POST"
    $request.ContentType   = "application/json; charset=utf-8"
    $request.Accept        = "application/json"
    $request.ContentLength = $requestBytes.Length
    $request.Timeout       = 30000

    try {
        $reqStream = $request.GetRequestStream()
        $reqStream.Write($requestBytes, 0, $requestBytes.Length)
        $reqStream.Close()

        $response   = $request.GetResponse()
        $reader     = [System.IO.StreamReader]::new($response.GetResponseStream(), [System.Text.Encoding]::UTF8)
        $rawContent = $reader.ReadToEnd()
        $reader.Close()
        $response.Close()
    }
    catch [System.Net.WebException] {
        $ex = $_.Exception
        $detail = if ($ex.Response) {
            $errReader = [System.IO.StreamReader]::new($ex.Response.GetResponseStream())
            $body = $errReader.ReadToEnd()
            $errReader.Close()
            $ex.Response.Close()
            $body
        } else { $ex.Message }

        Write-Warning "Jooble request failed for '$searchCountry' — writing empty output and continuing. ($detail)"
        [System.IO.File]::WriteAllText($currentOutputPath, "[]", $utf8NoBom)
        continue
    }
    catch {
        Write-Warning "Jooble request failed for '$searchCountry' — writing empty output and continuing. ($_)"
        [System.IO.File]::WriteAllText($currentOutputPath, "[]", $utf8NoBom)
        continue
    }

    # Parse the response
    $parsed = $null
    try {
        $parsed = $rawContent | ConvertFrom-Json
    }
    catch {
        Write-Warning "Jooble response for '$searchCountry' could not be parsed as JSON — writing empty output and continuing."
        [System.IO.File]::WriteAllText($currentOutputPath, "[]", $utf8NoBom)
        continue
    }

    if ($null -eq $parsed) {
        Write-Warning "Jooble API returned no data for '$searchCountry' — writing empty output and continuing."
        [System.IO.File]::WriteAllText($currentOutputPath, "[]", $utf8NoBom)
        continue
    }

    # Jooble returns { totalCount: N, jobs: [...] }
    $rawJobs = @(Get-Prop $parsed "jobs" @())

    if ($rawJobs.Count -eq 0) {
        Write-Warning "Jooble API returned no job records for '$searchCountry' — skipping file."
        continue
    }

    Write-Host "Received $($rawJobs.Count) jobs from Jooble for $searchCountry."

    # -----------------------------------------------------------------------
    # Transform
    # -----------------------------------------------------------------------

    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    $transformed = foreach ($job in $rawJobs) {

        # --- Core fields ---
        $id      = Normalize-String (Get-Prop $job "id")
        $title   = Normalize-String (Get-Prop $job "title")
        $link    = Normalize-String (Get-Prop $job "link")
        $snippet = Normalize-String (Get-Prop $job "snippet")
        $company = Normalize-String (Get-Prop $job "company")
        $type    = Normalize-String (Get-Prop $job "type")
        $locRaw  = Normalize-String (Get-Prop $job "location")
        $salary  = Normalize-String (Get-Prop $job "salary")
        $updated = Normalize-String (Get-Prop $job "updated")
        $image   = Normalize-String (Get-Prop $job "image")

        if ($null -eq $id)    { continue }
        if ($null -eq $title) { continue }
        if ($null -eq $link)  { continue }

        if (-not $seen.Add($id)) {
            Write-Warning "Duplicate Jooble id skipped for '$searchCountry': $id"
            continue
        }

        if ($null -eq $company) {
            $source = Normalize-String (Get-Prop $job "source")
            $company = if ($null -ne $source) { $source } else { "Unknown" }
        }

        $locationText  = $locRaw
        $workplaceType = Get-WorkplaceType -LocationText $locationText -Type $type
        $country       = Get-CountryFromLocation -LocationText $locationText
        $regions       = @(Get-RegionsFromLocationText -LocationText $locationText -Country $country)
        $countries     = if ($null -ne $country) { @($country) } else { @() }

        $empType = Convert-EmploymentType -Type $type -Title $title

        $snippetPlain = Strip-HtmlToPlainText -Html $snippet
        $category = Get-JobCategory -Title $title -SnippetPlain $snippetPlain

        $description = Build-Description -Snippet $snippet
        $summary     = Get-PlainTextSummary -Text $snippet -MaxLength 100

        $salaryParsed    = Parse-SalaryRange -SalaryText $salary
        $salaryFrom      = $salaryParsed.From
        $salaryTo        = $salaryParsed.To
        $salaryCurrency  = $salaryParsed.Currency

        $publishedUtc = ConvertTo-UtcIsoString $updated
        $slug         = ConvertTo-JobSlug -Title $title -Id $id

        $jobObj = [PSCustomObject][ordered]@{
            sourceSite             = "jooble.org"
            externalId             = $id
            companyName            = $company
            companyLogoUrl         = $image

            title                  = $title
            description            = $description
            workplaceType          = $workplaceType
            employmentType         = $empType

            externalApplicationUrl = $link

            category               = $category
            keywords               = $null
            regions                = @($regions)
            countries              = @($countries)

            summary                = $summary
            requirements           = $null
            benefits               = $null
            department             = $null

            salaryFrom             = $salaryFrom
            salaryTo               = $salaryTo
            salaryCurrency         = $salaryCurrency

            publishedUtc           = $publishedUtc
            postingExpiresUtc      = $null

            locationText           = $locationText
            locationCity           = $null
            locationState          = $null
            locationCountry        = $country
            locationCountryCode    = $null
            locationLatitude       = $null
            locationLongitude      = $null

            slug                   = $slug
        }

        if ($null -eq $jobObj.employmentType) {
            $jobObj.PSObject.Properties.Remove('employmentType')
        }

        if ($null -eq $jobObj.companyLogoUrl) {
            $jobObj.PSObject.Properties.Remove('companyLogoUrl')
        }

        $jobObj
    }

    $transformedArr = @($transformed | Where-Object { $null -ne $_ })

    if ($transformedArr.Count -eq 0) {
        Write-Warning "No valid Jooble records after transformation for '$searchCountry' — skipping file."
        continue
    }

    $jsonContent = ConvertTo-Json -InputObject $transformedArr -Depth 20 -AsArray

    try {
        [System.IO.File]::WriteAllText($currentOutputPath, $jsonContent, $utf8NoBom)
    }
    catch {
        throw "Failed to write output to '$currentOutputPath': $_"
    }

    Write-Host "Transformed job count : $($transformedArr.Count)"
    Write-Host "Output written to     : $currentOutputPath"
}
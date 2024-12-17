[cmdletbinding(DefaultParameterSetName=$false)]
Param(
  # Mandatory parameter set "Rule" with a list of country codes using ValidateSet for restriction
  [Parameter(ParameterSetName="Rule", Mandatory=$true)]
  [ValidateSet(
    "AD", "AE", "AF", "AG", "AI", "AL", "AM", "AO", "AQ", "AR", "AS", "AT", "AU", "AW", "AX", "AZ",
    "BA", "BB", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BL", "BM", "BN", "BO", "BQ", "BR", "BS",
    "BT", "BV", "BW", "BY", "BZ", "CA", "CC", "CD", "CF", "CG", "CH", "CI", "CK", "CL", "CM", "CN",
    "CO", "CR", "CU", "CV", "CW", "CX", "CY", "CZ", "DE", "DJ", "DK", "DM", "DO", "DZ", "EC", "EE",
    "EG", "EH", "ER", "ES", "ET", "FI", "FJ", "FK", "FM", "FO", "FR", "GA", "GB", "GD", "GE", "GF",
    "GG", "GH", "GI", "GL", "GM", "GN", "GP", "GQ", "GR", "GS", "GT", "GU", "GW", "GY", "HK", "HM",
    "HN", "HR", "HT", "HU", "ID", "IE", "IL", "IM", "IN", "IO", "IQ", "IR", "IS", "IT", "JE", "JM",
    "JO", "JP", "KE", "KG", "KH", "KI", "KM", "KN", "KP", "KR", "KW", "KY", "KZ", "LA", "LB", "LC",
    "LI", "LK", "LR", "LS", "LT", "LU", "LV", "LY", "MA", "MC", "MD", "ME", "MF", "MG", "MH", "MK",
    "ML", "MM", "MN", "MO", "MP", "MQ", "MR", "MS", "MT", "MU", "MV", "MW", "MX", "MY", "MZ", "NA",
    "NC", "NE", "NF", "NG", "NI", "NL", "NO", "NP", "NR", "NU", "NZ", "OM", "PA", "PE", "PF", "PG",
    "PH", "PK", "PL", "PM", "PN", "PR", "PS", "PT", "PW", "PY", "QA", "RE", "RO", "RS", "RU", "RW",
    "SA", "SB", "SC", "SD", "SE", "SG", "SH", "SI", "SJ", "SK", "SL", "SM", "SN", "SO", "SR", "SS",
    "ST", "SV", "SX", "SY", "SZ", "TC", "TD", "TF", "TG", "TH", "TJ", "TK", "TL", "TM", "TN", "TO",
    "TR", "TT", "TV", "TW", "TZ", "UA", "UG", "UM", "US", "UY", "UZ", "VA", "VC", "VE", "VG", "VI",
    "VN", "VU", "WF", "WS", "XK", "YE", "YT", "ZA", "ZM", "ZW"
  )]
  [string[]]$Country,

  # Optional parameters to target firewall rules
  [Parameter(ParameterSetName="Rule", Mandatory=$false)]
  [string[]]$RuleName,

  [Parameter(ParameterSetName="Rule", Mandatory=$false)]
  [string[]]$RuleDisplayName,

  [Parameter(ParameterSetName="Rule", Mandatory=$false)]
  [string[]]$RuleDisplayGroup,

  [Parameter(ParameterSetName="Rule", Mandatory=$false)]
  [switch]$ExcludeLocalSubnet, # Currently not used

  # Parameter set to list countries
  [Parameter(ParameterSetName="ListCountries", Mandatory=$true)]
  [switch]$ListCountries,

  # MaxMind license key for downloading GeoIP data
  [Parameter(Mandatory=$false)]
  [string]$MaxMindLicenseKey,

  # Force re-download of the GeoIP database
  [Parameter(Mandatory=$false)]
  [switch]$ForceDownload=$false
)

# Stop execution on errors
$ErrorActionPreference = "Stop"

# Constants
$APP_DIR = Join-Path $env:LOCALAPPDATA "GeoWF"
$MAXMIND_LICENSE_KEY_FILE = Join-Path $APP_DIR "maxmind_license_key.txt"
$GEOIP_URL_TEMPLATE = "https://download.maxmind.com/app/geoip_download?edition_id=GeoLite2-Country-CSV&license_key={0}&suffix=zip"

# Maximum number of ranges per firewall rule
$MAXIMUM_RANGES = 10000

# Determine if we are updating rules
$RuleUpdateDesired = $PSCmdlet.ParameterSetName -eq "Rule"

# Create the application directory if it doesn't exist
if (!(Test-Path $APP_DIR)) {
  [void](New-Item -ItemType Directory -Force $APP_DIR)
}

# Save or load the MaxMind license key
if ($MaxMindLicenseKey) {
  ($MaxMindLicenseKey -replace "\s", "") | Out-File -FilePath $MAXMIND_LICENSE_KEY_FILE -NoNewline
  Write-Information ("License key saved to `"{0}`"." -f $MAXMIND_LICENSE_KEY_FILE)
}
if (!(Test-Path $MAXMIND_LICENSE_KEY_FILE)) {
  throw "License key file not found. Set -MaxMindLicenseKey parameter to save a valid license key. Create a free account at https://www.maxmind.com/en/geolite2/signup."
}
$MaxMindLicenseKey = (Get-Content $MAXMIND_LICENSE_KEY_FILE) -replace "\s", ""
if (!$RuleUpdateDesired -and !$ForceDownload -and !$ListCountries) { exit }

# Define GeoIP paths
$GeoIPURL = ($GEOIP_URL_TEMPLATE -f $MaxMindLicenseKey)
$GeoIPDir = Join-Path $APP_DIR "GeoIP"
$GeoIPZip = Join-Path $APP_DIR "GeoIP.zip"

# Re-download GeoIP data if forced
if ($ForceDownload -and (Test-Path $GeoIPDir)) {
  Write-Information "Deleting existing GeoIP data."
  Remove-Item $GeoIPDir -Recurse -Force -Confirm:$false
}

# Download and extract GeoIP database
if (!(Test-Path $GeoIPDir)) {
  Write-Information "Downloading GeoIP database..."
  Invoke-WebRequest $GeoIPURL -OutFile $GeoIPZip
  Write-Information "Extracting GeoIP database..."
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  [System.IO.Compression.ZipFile]::ExtractToDirectory($GeoIPZip, $APP_DIR)
  Rename-Item -Path (Get-Item -Path (Join-Path $APP_DIR "GeoLite2-Country-CSV_*") | Sort-Object Name -Descending)[0] -NewName "GeoIP"
  Remove-Item $GeoIPZip -Force -Confirm:$false
}

# Import country and network data
Write-Information "Importing country data..."
$CountryLocations = Import-Csv (Join-Path $GeoIPDir "GeoLite2-Country-Locations-en.csv")

if ($ListCountries) {
  $CountryLocations | Where-Object { $_.country_name -match "\w" } | Sort-Object country_name | Select-Object `
    @{ n = "CountryName"; e = { $_.country_name } },
    @{ n = "CountryCode"; e = { $_.country_iso_code } }
  exit
}

Write-Information "Importing IPv4 network blocks..."
$CountryBlocksIPv4 = Import-Csv (Join-Path $GeoIPDir "GeoLite2-Country-Blocks-IPv4.csv")

# Match countries to GeoNames IDs
$CountriesGeonameIDs = ($CountryLocations | Where-Object { $Country -contains $_.country_iso_code }).geoname_id

# Extract IP networks matching the selected countries
Write-Information "Extracting IP networks..."
$Networks = ($CountryBlocksIPv4 | Where-Object { $CountriesGeonameIDs -contains $_.geoname_id }).network

# Search for matching firewall rules
$TargetRules = @()
if ($RuleName) { $TargetRules += Get-NetFirewallRule -Name $RuleName | Where-Object { $_.Direction -eq "Inbound" } }
if ($RuleDisplayName) { $TargetRules += Get-NetFirewallRule -DisplayName $RuleDisplayName | Where-Object { $_.Direction -eq "Inbound" } }
if ($RuleDisplayGroup) {
  try {
    $TargetRules += Get-NetFirewallRule -DisplayGroup $RuleDisplayGroup | Where-Object { $_.Direction -eq "Inbound" }
  } catch {
    Write-Information "No rules found for DisplayGroup '$RuleDisplayGroup'. New rules will be created."
  }
}

if ($TargetRules.Count -gt 0) {
    Write-Host "Updating existing firewall rules with IP ranges."

    $NumberOfChunks = [math]::Ceiling($Networks.Count / $MAXIMUM_RANGES)

    for ($chunkIndex = 0; $chunkIndex -lt $NumberOfChunks; $chunkIndex++) {
        $start = $chunkIndex * $MAXIMUM_RANGES
        $end = [math]::Min($start + $MAXIMUM_RANGES - 1, $Networks.Count - 1)
        $currentChunk = $Networks[$start..$end]

        Set-NetFirewallRule -Name $TargetRules.Name -RemoteAddress $currentChunk
    }
} else {
    Write-Warning "No matching firewall rules found. Creating new rules."

    $NumberOfChunks = [math]::Ceiling($Networks.Count / $MAXIMUM_RANGES)

    for ($chunkIndex = 0; $chunkIndex -lt $NumberOfChunks; $chunkIndex++) {
        $start = $chunkIndex * $MAXIMUM_RANGES
        $end = [math]::Min($start + $MAXIMUM_RANGES - 1, $Networks.Count - 1)
        $currentChunk = $Networks[$start..$end]

        $newRuleName = "GeoRule_$RuleDisplayGroup_Part$($chunkIndex+1)"
        $newDisplayName = "$RuleDisplayGroup (Part $($chunkIndex+1))"

        Write-Host "Creating new rule: $newRuleName"

        # Ensure -Group is always a string
        New-NetFirewallRule -Name $newRuleName `
                            -DisplayName $newDisplayName `
                            -Direction Inbound `
                            -Action Block `
                            -RemoteAddress $currentChunk `
                            -Group ($RuleDisplayGroup -join "") `
                            -Protocol Any `
                            -Verbose
    }
}

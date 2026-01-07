<#
.SYNOPSIS
Shared helper functions for SandboxStart

.DESCRIPTION
Provides utility functions including:
- GitHub API interaction with intelligent caching and fallback mechanisms
- Rate limiting handling
- Fallback to Atom feeds when API fails

.NOTES
Author: SandboxStart
Version: 1.0.0
#>

# Script-level variables
$script:CacheDirectory = Join-Path $env:LOCALAPPDATA "Packages\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe\SandboxTest\GitHubCache"
$script:CacheMetadataFile = Join-Path $script:CacheDirectory "cache_metadata.json"

#region Cache Management Functions

function Initialize-CacheDirectory {
	<#
	.SYNOPSIS
	Ensures cache directory exists
	#>
	if (!(Test-Path $script:CacheDirectory)) {
		New-Item -Path $script:CacheDirectory -ItemType Directory -Force | Out-Null
		Write-Verbose "Created GitHub cache directory: $script:CacheDirectory"
	}
}

function Get-CacheKey {
	<#
	.SYNOPSIS
	Generates a cache key from URI

	.PARAMETER Uri
	GitHub API URI

	.OUTPUTS
	String cache key
	#>
	param(
		[Parameter(Mandatory=$true)]
		[string]$Uri
	)

	# Convert URI to a safe filename
	$key = $Uri -replace 'https://api\.github\.com/repos/', '' `
				-replace 'https://api\.github\.com/', '' `
				-replace '[/?&=]', '_'

	return $key
}

function Get-CacheFilePath {
	<#
	.SYNOPSIS
	Gets the cache file path for a given URI

	.PARAMETER Uri
	GitHub API URI

	.OUTPUTS
	String file path
	#>
	param(
		[Parameter(Mandatory=$true)]
		[string]$Uri
	)

	$key = Get-CacheKey -Uri $Uri
	return Join-Path $script:CacheDirectory "$key.json"
}

function Get-CacheMetadata {
	<#
	.SYNOPSIS
	Loads cache metadata from disk

	.OUTPUTS
	Hashtable of metadata
	#>
	if (!(Test-Path $script:CacheMetadataFile)) {
		return @{}
	}

	try {
		$json = Get-Content $script:CacheMetadataFile -Raw -ErrorAction Stop
		$obj = $json | ConvertFrom-Json

		# Convert PSCustomObject to Hashtable (PowerShell 5.1 compatibility)
		$hashtable = @{}
		$obj.PSObject.Properties | ForEach-Object {
			$hashtable[$_.Name] = $_.Value
		}

		return $hashtable
	}
	catch {
		Write-Warning "Failed to load cache metadata: $_"
		return @{}
	}
}

function Save-CacheMetadata {
	<#
	.SYNOPSIS
	Saves cache metadata to disk

	.PARAMETER Metadata
	Hashtable of metadata
	#>
	param(
		[Parameter(Mandatory=$true)]
		[hashtable]$Metadata
	)

	try {
		Initialize-CacheDirectory
		$Metadata | ConvertTo-Json -Depth 10 | Set-Content $script:CacheMetadataFile -Force
	}
	catch {
		Write-Warning "Failed to save cache metadata: $_"
	}
}

function Test-CacheExpired {
	<#
	.SYNOPSIS
	Checks if cache entry is expired

	.PARAMETER CacheKey
	Cache key to check

	.PARAMETER Metadata
	Cache metadata hashtable

	.OUTPUTS
	Boolean indicating if cache is expired
	#>
	param(
		[Parameter(Mandatory=$true)]
		[string]$CacheKey,

		[Parameter(Mandatory=$true)]
		[hashtable]$Metadata
	)

	if (!$Metadata.ContainsKey($CacheKey)) {
		return $true
	}

	$entry = $Metadata[$CacheKey]
	if (!$entry.expires_at) {
		return $true
	}

	try {
		$expiresAt = [DateTime]::Parse($entry.expires_at)
		return ([DateTime]::UtcNow -gt $expiresAt)
	}
	catch {
		return $true
	}
}

function Get-CachedResponse {
	<#
	.SYNOPSIS
	Retrieves cached response if available

	.PARAMETER Uri
	GitHub API URI

	.PARAMETER IgnoreExpiry
	Return cached data even if expired

	.OUTPUTS
	Cached data or $null
	#>
	param(
		[Parameter(Mandatory=$true)]
		[string]$Uri,

		[switch]$IgnoreExpiry
	)

	$cacheKey = Get-CacheKey -Uri $Uri
	$cacheFile = Get-CacheFilePath -Uri $Uri

	if (!(Test-Path $cacheFile)) {
		return $null
	}

	$metadata = Get-CacheMetadata

	if (!$IgnoreExpiry -and (Test-CacheExpired -CacheKey $cacheKey -Metadata $metadata)) {
		Write-Verbose "Cache expired for: $cacheKey"
		return $null
	}

	try {
		$json = Get-Content $cacheFile -Raw -ErrorAction Stop
		$data = $json | ConvertFrom-Json
		Write-Verbose "Cache hit for: $cacheKey"
		return $data
	}
	catch {
		Write-Warning "Failed to load cached data: $_"
		return $null
	}
}

function Save-CachedResponse {
	<#
	.SYNOPSIS
	Saves API response to cache

	.PARAMETER Uri
	GitHub API URI

	.PARAMETER Data
	Response data to cache

	.PARAMETER TTL
	Time-to-live in minutes

	.PARAMETER Source
	Source of data (api, atom, etc.)
	#>
	param(
		[Parameter(Mandatory=$true)]
		[string]$Uri,

		[Parameter(Mandatory=$true)]
		$Data,

		[Parameter(Mandatory=$false)]
		[int]$TTL = 60,

		[Parameter(Mandatory=$false)]
		[string]$Source = 'api'
	)

	try {
		Initialize-CacheDirectory

		$cacheKey = Get-CacheKey -Uri $Uri
		$cacheFile = Get-CacheFilePath -Uri $Uri

		# Save data
		$Data | ConvertTo-Json -Depth 100 | Set-Content $cacheFile -Force

		# Update metadata
		$metadata = Get-CacheMetadata
		$timestamp = [DateTime]::UtcNow
		$expiresAt = $timestamp.AddMinutes($TTL)

		$metadata[$cacheKey] = @{
			timestamp = $timestamp.ToString('o')
			ttl_minutes = $TTL
			expires_at = $expiresAt.ToString('o')
			source = $Source
		}

		Save-CacheMetadata -Metadata $metadata
		Write-Verbose "Cached response for: $cacheKey (TTL: $TTL min, Source: $Source)"
	}
	catch {
		Write-Warning "Failed to cache response: $_"
	}
}

function Clear-GitHubCache {
	<#
	.SYNOPSIS
	Clears all cached GitHub API data

	.DESCRIPTION
	Removes the entire GitHub cache directory and all cached responses.
	This function is called when the user selects the "Clean (cached dependencies)" option.
	#>
	if (Test-Path $script:CacheDirectory) {
		try {
			Remove-Item -Path $script:CacheDirectory -Recurse -Force -ErrorAction Stop
			Write-Verbose "GitHub cache cleared successfully"
		}
		catch {
			Write-Warning "Failed to clear GitHub cache: $_"
		}
	}
	else {
		Write-Verbose "GitHub cache directory does not exist, nothing to clear"
	}
}

#endregion

#region GitHub API Functions

function Get-GitHubPersonalAccessToken {
	<#
	.SYNOPSIS
	Retrieves GitHub Personal Access Token from environment variable

	.DESCRIPTION
	Checks for $env:GITHUB_PAT environment variable. This is only needed
	for development/testing by repository owners. End users do not need PAT.

	.OUTPUTS
	String (PAT) or $null if not configured
	#>
	if ($env:GITHUB_PAT) {
		Write-Verbose "GitHub PAT found in environment"
		return $env:GITHUB_PAT
	}

	return $null
}

function Test-GitHubApiLimit {
	<#
	.SYNOPSIS
	Checks GitHub API rate limit status

	.OUTPUTS
	PSCustomObject with rate limit information
	#>
	try {
		$headers = @{
			'User-Agent' = 'SandboxStart-PowerShell'
		}

		$pat = Get-GitHubPersonalAccessToken
		if ($pat) {
			$headers['Authorization'] = "Bearer $pat"
		}

		$response = Invoke-RestMethod -Uri 'https://api.github.com/rate_limit' `
			-Headers $headers -TimeoutSec 5 -UseBasicParsing -ErrorAction Stop

		return [PSCustomObject]@{
			Remaining = $response.rate.remaining
			Limit = $response.rate.limit
			ResetTime = [DateTimeOffset]::FromUnixTimeSeconds($response.rate.reset).LocalDateTime
		}
	}
	catch {
		Write-Warning "Failed to check API rate limit: $_"
		return $null
	}
}

function Invoke-GitHubApi {
	<#
	.SYNOPSIS
	Wrapper for GitHub API calls with caching and fallback

	.PARAMETER Uri
	GitHub API endpoint URI

	.PARAMETER UseCache
	Use cached response if available

	.PARAMETER CacheTTL
	Cache time-to-live in minutes (default: 60)

	.PARAMETER FallbackToFeed
	Use Atom feed if API fails

	.OUTPUTS
	API response data
	#>
	param(
		[Parameter(Mandatory=$true)]
		[string]$Uri,

		[switch]$UseCache,

		[int]$CacheTTL = 60,

		[switch]$FallbackToFeed
	)

	# Check cache first
	if ($UseCache) {
		$cached = Get-CachedResponse -Uri $Uri
		if ($cached) {
			return $cached
		}
	}

	# Prepare headers
	$headers = @{
		'User-Agent' = 'SandboxStart-PowerShell'
	}

	$pat = Get-GitHubPersonalAccessToken
	if ($pat) {
		$headers['Authorization'] = "Bearer $pat"
		Write-Verbose "Using authenticated API request (PAT)"
	}

	# Make API request
	try {
		Write-Verbose "Fetching from GitHub API: $Uri"
		$response = Invoke-RestMethod -Uri $Uri -Headers $headers -TimeoutSec 5 -UseBasicParsing -ErrorAction Stop

		# Cache successful response
		if ($UseCache) {
			Save-CachedResponse -Uri $Uri -Data $response -TTL $CacheTTL -Source 'api'
		}

		return $response
	}
	catch [Microsoft.PowerShell.Commands.HttpResponseException] {
		$statusCode = $_.Exception.Response.StatusCode.value__

		if ($statusCode -eq 403) {
			Write-Warning "GitHub API rate limit exceeded (403). Consider setting `$env:GITHUB_PAT for higher limits (5,000/hour vs 60/hour)"

			# Try to use expired cache as fallback
			if ($UseCache) {
				$cached = Get-CachedResponse -Uri $Uri -IgnoreExpiry
				if ($cached) {
					Write-Warning "Using expired cache data as fallback"
					return $cached
				}
			}

			# Try Atom feed fallback for releases endpoints
			if ($FallbackToFeed -and $Uri -match '/releases') {
				Write-Verbose "Attempting Atom feed fallback"
				$feedUrl = $Uri -replace 'api\.github\.com/repos/(.*?)/releases.*', 'github.com/$1/releases.atom'
				$feedData = ConvertFrom-AtomFeed -FeedUrl $feedUrl

				if ($feedData) {
					if ($UseCache) {
						Save-CachedResponse -Uri $Uri -Data $feedData -TTL $CacheTTL -Source 'atom'
					}
					return $feedData
				}
			}
		}

		throw
	}
}

function ConvertFrom-AtomFeed {
	<#
	.SYNOPSIS
	Parses GitHub Atom feed as fallback for releases

	.PARAMETER FeedUrl
	Atom feed URL (https://github.com/{owner}/{repo}/releases.atom)

	.OUTPUTS
	Array of release objects compatible with API format
	#>
	param(
		[Parameter(Mandatory=$true)]
		[string]$FeedUrl
	)

	try {
		Write-Verbose "Fetching Atom feed: $FeedUrl"
		[xml]$feed = (Invoke-WebRequest -Uri $FeedUrl -UseBasicParsing -TimeoutSec 5).Content

		if (!$feed) {
			Write-Warning "Failed to parse Atom feed"
			return @()
		}

		# Define XML namespace
		$ns = @{atom='http://www.w3.org/2005/Atom'}

		$releases = @()
		$entries = $feed.SelectNodes('//atom:entry', (New-Object System.Xml.XmlNamespaceManager($feed.NameTable)) -replace '', $ns)

		# Fallback if SelectNodes doesn't work
		if (!$entries -and $feed.feed.entry) {
			$entries = $feed.feed.entry
		}

		foreach ($entry in $entries) {
			$title = if ($entry.title.'#text') { $entry.title.'#text' } else { $entry.title }
			$published = if ($entry.published.'#text') { $entry.published.'#text' } else { $entry.published }

			# Extract tag from title (e.g., "v1.7.10514")
			$tag = $null
			if ($title -match 'v?\d+\.\d+\.\d+') {
				$tag = $matches[0]
			}

			if ($tag) {
				$releases += [PSCustomObject]@{
					tag_name = $tag
					published_at = $published
					prerelease = ($title -match 'pre-?release|preview|beta|alpha')
					assets = @()  # Limited in Atom feed
					source = 'atom'
				}
			}
		}

		Write-Verbose "Parsed $($releases.Count) releases from Atom feed"
		return $releases
	}
	catch {
		Write-Warning "Failed to parse Atom feed: $_"
		return @()
	}
}

function Get-GitHubReleases {
	<#
	.SYNOPSIS
	Get releases for a repository with intelligent fallback

	.PARAMETER Owner
	Repository owner (e.g., "microsoft")

	.PARAMETER Repo
	Repository name (e.g., "winget-cli")

	.PARAMETER PerPage
	Number of releases to fetch (max 100)

	.PARAMETER StableOnly
	Filter out prereleases

	.PARAMETER UseCache
	Use cached response if available

	.OUTPUTS
	Array of release objects with: tag_name, prerelease, assets[], published_at
	#>
	param(
		[Parameter(Mandatory=$true)]
		[string]$Owner,

		[Parameter(Mandatory=$true)]
		[string]$Repo,

		[int]$PerPage = 100,

		[switch]$StableOnly,

		[switch]$UseCache
	)

	$uri = "https://api.github.com/repos/$Owner/$Repo/releases?per_page=$PerPage"

	$releases = Invoke-GitHubApi -Uri $uri -UseCache:$UseCache -FallbackToFeed

	if ($StableOnly) {
		$releases = $releases | Where-Object { -not $_.prerelease }
	}

	return $releases
}

function Get-GitHubFolderContents {
	<#
	.SYNOPSIS
	Get folder contents with caching

	.PARAMETER Owner
	Repository owner

	.PARAMETER Repo
	Repository name

	.PARAMETER Path
	Folder path in repository

	.PARAMETER Branch
	Branch name (default: "master")

	.PARAMETER FilePattern
	Filter files by pattern (e.g., "*.ps1")

	.PARAMETER UseCache
	Use cached response if available

	.OUTPUTS
	Array of file objects with: name, type, download_url
	#>
	param(
		[Parameter(Mandatory=$true)]
		[string]$Owner,

		[Parameter(Mandatory=$true)]
		[string]$Repo,

		[Parameter(Mandatory=$true)]
		[string]$Path,

		[string]$Branch = 'master',

		[string]$FilePattern,

		[switch]$UseCache
	)

	$uri = "https://api.github.com/repos/$Owner/$Repo/contents/$Path`?ref=$Branch"

	$files = Invoke-GitHubApi -Uri $uri -UseCache:$UseCache

	if ($FilePattern) {
		$files = $files | Where-Object {
			$_.type -eq 'file' -and $_.name -like $FilePattern
		}
	}

	return $files
}

function Get-GitHubLatestRelease {
	<#
	.SYNOPSIS
	Get latest release info with Atom feed fallback

	.PARAMETER Owner
	Repository owner

	.PARAMETER Repo
	Repository name

	.PARAMETER UseCache
	Use cached response if available

	.OUTPUTS
	Release object with: tag_name, assets[], created_at
	#>
	param(
		[Parameter(Mandatory=$true)]
		[string]$Owner,

		[Parameter(Mandatory=$true)]
		[string]$Repo,

		[switch]$UseCache
	)

	$uri = "https://api.github.com/repos/$Owner/$Repo/releases/latest"

	return Invoke-GitHubApi -Uri $uri -UseCache:$UseCache -FallbackToFeed
}

#endregion

#region Windows Theme Detection

function Get-WindowsThemeSetting {
	<#
	.SYNOPSIS
	Detects Windows dark/light mode preference from registry

	.DESCRIPTION
	Reads the AppsUseLightTheme registry value to determine if dark mode is enabled.
	Returns $true if dark mode should be used, $false for light mode.

	.OUTPUTS
	Boolean - $true for dark mode, $false for light mode

	.EXAMPLE
	$isDark = Get-WindowsThemeSetting
	if ($isDark) { Write-Host "Dark mode enabled" }
	#>

	try {
		$appsUseLightTheme = (Get-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "AppsUseLightTheme" -ErrorAction SilentlyContinue).AppsUseLightTheme
		# If value is 0 or cannot be read, use dark mode
		return ($null -eq $appsUseLightTheme -or $appsUseLightTheme -eq 0)
	}
	catch {
		# Default to dark mode if registry cannot be read
		return $true
	}
}

#endregion

#region SandboxStart Theme Preferences

function global:Get-SandboxStartThemePreference {
	<#
	.SYNOPSIS
	Reads SandboxStart theme preference from registry

	.DESCRIPTION
	Retrieves the user's saved theme mode from registry.
	Defaults to "Auto" (follow Windows system theme) if not set.

	.OUTPUTS
	String - "Auto", "Light", "Dark", or "Custom"

	.EXAMPLE
	$themeMode = Get-SandboxStartThemePreference
	if ($themeMode -eq "Dark") { Write-Host "Dark theme selected" }
	#>

	try {
		$regPath = "HKCU:\Software\SandboxStart"
		$value = (Get-ItemProperty -Path $regPath -Name "ThemeMode" -ErrorAction SilentlyContinue).ThemeMode

		if ($null -eq $value -or $value -notin @("Auto", "Light", "Dark", "Custom")) {
			return "Auto"
		}

		return $value
	}
	catch {
		return "Auto"
	}
}

function global:Set-SandboxStartThemePreference {
	<#
	.SYNOPSIS
	Saves SandboxStart theme preference to registry

	.DESCRIPTION
	Stores the user's theme mode selection in registry for persistence across sessions.

	.PARAMETER ThemeMode
	The theme mode to save: "Auto", "Light", "Dark", or "Custom"

	.EXAMPLE
	Set-SandboxStartThemePreference -ThemeMode "Dark"
	#>

	param(
		[Parameter(Mandatory = $true)]
		[ValidateSet("Auto", "Light", "Dark", "Custom")]
		[string]$ThemeMode
	)

	try {
		$regPath = "HKCU:\Software\SandboxStart"

		# Create registry key if it doesn't exist
		if (-not (Test-Path $regPath)) {
			New-Item -Path $regPath -Force | Out-Null
		}

		# Save theme mode
		Set-ItemProperty -Path $regPath -Name "ThemeMode" -Value $ThemeMode -Type String
	}
	catch {
		Write-Warning "Failed to save theme preference: $($_.Exception.Message)"
	}
}

function global:Get-SandboxStartCustomColors {
	<#
	.SYNOPSIS
	Reads custom color settings from registry

	.DESCRIPTION
	Retrieves user's custom color configuration for all 6 theme elements.
	Returns default dark mode colors if not set.

	.OUTPUTS
	Hashtable with 6 color elements as RGB strings ("R,G,B")

	.EXAMPLE
	$colors = Get-SandboxStartCustomColors
	Write-Host "Background: $($colors.BackColor)"
	#>

	try {
		$regPath = "HKCU:\Software\SandboxStart\CustomColors"

		# Default dark mode colors
		$defaults = @{
			BackColor         = "32,32,32"
			ForeColor         = "255,255,255"
			ButtonBackColor   = "70,70,70"
			TextBoxBackColor  = "45,45,45"
			GrayLabelColor    = "180,180,180"
			UpdateButtonColor = "60,120,60"
		}

		if (-not (Test-Path $regPath)) {
			return $defaults
		}

		$props = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue

		# Build hashtable, using defaults for missing values
		$colors = @{
			BackColor         = if ($props.BackColor) { $props.BackColor } else { $defaults.BackColor }
			ForeColor         = if ($props.ForeColor) { $props.ForeColor } else { $defaults.ForeColor }
			ButtonBackColor   = if ($props.ButtonBackColor) { $props.ButtonBackColor } else { $defaults.ButtonBackColor }
			TextBoxBackColor  = if ($props.TextBoxBackColor) { $props.TextBoxBackColor } else { $defaults.TextBoxBackColor }
			GrayLabelColor    = if ($props.GrayLabelColor) { $props.GrayLabelColor } else { $defaults.GrayLabelColor }
			UpdateButtonColor = if ($props.UpdateButtonColor) { $props.UpdateButtonColor } else { $defaults.UpdateButtonColor }
		}

		return $colors
	}
	catch {
		Write-Warning "Failed to read custom colors: $($_.Exception.Message)"
		# Return default dark mode colors on error
		return @{
			BackColor         = "32,32,32"
			ForeColor         = "255,255,255"
			ButtonBackColor   = "70,70,70"
			TextBoxBackColor  = "45,45,45"
			GrayLabelColor    = "180,180,180"
			UpdateButtonColor = "60,120,60"
		}
	}
}

function global:Set-SandboxStartCustomColors {
	<#
	.SYNOPSIS
	Saves custom color settings to registry

	.DESCRIPTION
	Stores user's custom color configuration for all 6 theme elements.

	.PARAMETER Colors
	Hashtable containing all 6 required color elements as RGB strings ("R,G,B")

	.EXAMPLE
	$colors = @{
		BackColor = "32,32,32"
		ForeColor = "255,255,255"
		ButtonBackColor = "70,70,70"
		TextBoxBackColor = "45,45,45"
		GrayLabelColor = "180,180,180"
		UpdateButtonColor = "60,120,60"
	}
	Set-SandboxStartCustomColors -Colors $colors
	#>

	param(
		[Parameter(Mandatory = $true)]
		[hashtable]$Colors
	)

	try {
		$regPath = "HKCU:\Software\SandboxStart\CustomColors"

		# Validate hashtable contains all required keys
		$requiredKeys = @("BackColor", "ForeColor", "ButtonBackColor", "TextBoxBackColor", "GrayLabelColor", "UpdateButtonColor")
		foreach ($key in $requiredKeys) {
			if (-not $Colors.ContainsKey($key)) {
				throw "Missing required color element: $key"
			}
		}

		# Create registry key if it doesn't exist
		if (-not (Test-Path $regPath)) {
			New-Item -Path $regPath -Force | Out-Null
		}

		# Save each color element
		Set-ItemProperty -Path $regPath -Name "BackColor" -Value $Colors.BackColor -Type String
		Set-ItemProperty -Path $regPath -Name "ForeColor" -Value $Colors.ForeColor -Type String
		Set-ItemProperty -Path $regPath -Name "ButtonBackColor" -Value $Colors.ButtonBackColor -Type String
		Set-ItemProperty -Path $regPath -Name "TextBoxBackColor" -Value $Colors.TextBoxBackColor -Type String
		Set-ItemProperty -Path $regPath -Name "GrayLabelColor" -Value $Colors.GrayLabelColor -Type String
		Set-ItemProperty -Path $regPath -Name "UpdateButtonColor" -Value $Colors.UpdateButtonColor -Type String
	}
	catch {
		Write-Warning "Failed to save custom colors: $($_.Exception.Message)"
	}
}

function global:Export-SandboxStartTheme {
	<#
	.SYNOPSIS
	Exports a custom theme to a JSON file

	.DESCRIPTION
	Saves custom theme colors to a JSON file in the themes directory.
	Creates the themes directory if it doesn't exist.

	.PARAMETER Colors
	Hashtable containing the 6 color elements (BackColor, ForeColor, ButtonBackColor, TextBoxBackColor, GrayLabelColor, UpdateButtonColor)

	.PARAMETER ThemeName
	Name of the theme (used as filename)

	.PARAMETER FilePath
	Optional. Full path to save file. If not specified, saves to "$PSScriptRoot\..\themes\$ThemeName.json"

	.EXAMPLE
	Export-SandboxStartTheme -Colors $myColors -ThemeName "My Matrix Theme"

	.OUTPUTS
	Returns the full path to the saved file
	#>
	param(
		[Parameter(Mandatory = $true)]
		[hashtable]$Colors,

		[Parameter(Mandatory = $true)]
		[string]$ThemeName,

		[Parameter(Mandatory = $false)]
		[string]$FilePath
	)

	try {
		# Determine save path
		if (-not $FilePath) {
			$themesDir = Join-Path $PSScriptRoot "..\themes"
			if (-not (Test-Path $themesDir)) {
				New-Item -ItemType Directory -Path $themesDir -Force | Out-Null
			}

			# Sanitize theme name for filename
			$safeThemeName = $ThemeName -replace '[\\/:*?"<>|]', '_'
			$FilePath = Join-Path $themesDir "$safeThemeName.json"
		}

		# Create theme object
		$themeObject = @{
			ThemeName   = $ThemeName
			CreatedDate = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
			Colors      = @{
				BackColor         = $Colors.BackColor
				ForeColor         = $Colors.ForeColor
				ButtonBackColor   = $Colors.ButtonBackColor
				TextBoxBackColor  = $Colors.TextBoxBackColor
				GrayLabelColor    = $Colors.GrayLabelColor
				UpdateButtonColor = $Colors.UpdateButtonColor
			}
		}

		# Export to JSON
		$themeObject | ConvertTo-Json -Depth 3 | Out-File -FilePath $FilePath -Encoding UTF8

		return $FilePath
	}
	catch {
		Write-Error "Failed to export theme: $($_.Exception.Message)"
		return $null
	}
}

function global:Import-SandboxStartTheme {
	<#
	.SYNOPSIS
	Imports a custom theme from a JSON file

	.DESCRIPTION
	Loads custom theme colors from a JSON file and validates the structure.

	.PARAMETER FilePath
	Full path to the JSON theme file

	.EXAMPLE
	$theme = Import-SandboxStartTheme -FilePath "C:\themes\MyTheme.json"

	.OUTPUTS
	Returns hashtable with ThemeName and Colors properties, or $null on error
	#>
	param(
		[Parameter(Mandatory = $true)]
		[string]$FilePath
	)

	try {
		# Check if file exists
		if (-not (Test-Path $FilePath)) {
			Write-Error "Theme file not found: $FilePath"
			return $null
		}

		# Read and parse JSON
		$json = Get-Content -Path $FilePath -Raw -Encoding UTF8
		$themeObject = $json | ConvertFrom-Json

		# Validate required properties
		if (-not $themeObject.Colors) {
			Write-Error "Invalid theme file: Missing Colors property"
			return $null
		}

		$requiredKeys = @("BackColor", "ForeColor", "ButtonBackColor", "TextBoxBackColor", "GrayLabelColor", "UpdateButtonColor")
		$missingKeys = @()

		foreach ($key in $requiredKeys) {
			if (-not $themeObject.Colors.PSObject.Properties[$key]) {
				$missingKeys += $key
			}
		}

		if ($missingKeys.Count -gt 0) {
			Write-Error "Invalid theme file: Missing color keys: $($missingKeys -join ', ')"
			return $null
		}

		# Return as hashtable
		return @{
			ThemeName = if ($themeObject.ThemeName) { $themeObject.ThemeName } else { "Imported Theme" }
			Colors    = @{
				BackColor         = $themeObject.Colors.BackColor
				ForeColor         = $themeObject.Colors.ForeColor
				ButtonBackColor   = $themeObject.Colors.ButtonBackColor
				TextBoxBackColor  = $themeObject.Colors.TextBoxBackColor
				GrayLabelColor    = $themeObject.Colors.GrayLabelColor
				UpdateButtonColor = $themeObject.Colors.UpdateButtonColor
			}
		}
	}
	catch {
		Write-Error "Failed to import theme: $($_.Exception.Message)"
		return $null
	}
}

function global:Get-SandboxStartThemeFiles {
	<#
	.SYNOPSIS
	Gets list of available theme files

	.DESCRIPTION
	Scans the themes directory for JSON theme files and returns their information.

	.PARAMETER ThemesPath
	Optional. Path to themes directory. Defaults to "$PSScriptRoot\..\themes"

	.EXAMPLE
	$themes = Get-SandboxStartThemeFiles

	.OUTPUTS
	Returns array of hashtables with Name, Path, and CreatedDate properties
	#>
	param(
		[Parameter(Mandatory = $false)]
		[string]$ThemesPath
	)

	try {
		if (-not $ThemesPath) {
			$ThemesPath = Join-Path $PSScriptRoot "..\themes"
		}

		if (-not (Test-Path $ThemesPath)) {
			return @()
		}

		$themeFiles = Get-ChildItem -Path $ThemesPath -Filter "*.json" -File

		$themes = @()
		foreach ($file in $themeFiles) {
			try {
				$json = Get-Content -Path $file.FullName -Raw -Encoding UTF8
				$themeObject = $json | ConvertFrom-Json

				$themes += @{
					Name        = if ($themeObject.ThemeName) { $themeObject.ThemeName } else { $file.BaseName }
					Path        = $file.FullName
					CreatedDate = if ($themeObject.CreatedDate) { $themeObject.CreatedDate } else { $file.LastWriteTime.ToString("yyyy-MM-dd") }
				}
			}
			catch {
				# Skip invalid theme files
				Write-Warning "Skipping invalid theme file: $($file.Name)"
			}
		}

		return $themes
	}
	catch {
		Write-Warning "Failed to get theme files: $($_.Exception.Message)"
		return @()
	}
}

function global:Test-ColorIsDark {
	<#
	.SYNOPSIS
	Determines if a color is dark or light based on perceived brightness

	.DESCRIPTION
	Calculates perceived brightness using the luminance formula (0.299*R + 0.587*G + 0.114*B).
	Returns $true if the color is dark (brightness < 128), $false if light.
	Used to determine appropriate title bar theme for custom colors.

	.PARAMETER Color
	System.Drawing.Color object to test

	.OUTPUTS
	Boolean - $true if dark, $false if light

	.EXAMPLE
	$color = [System.Drawing.Color]::FromArgb(32, 32, 32)
	if (Test-ColorIsDark -Color $color) { Write-Host "This is a dark color" }
	#>

	param(
		[Parameter(Mandatory = $true)]
		[System.Drawing.Color]$Color
	)

	# Calculate perceived brightness using luminance formula
	$brightness = (0.299 * $Color.R + 0.587 * $Color.G + 0.114 * $Color.B)
	return $brightness -lt 128
}

function Sync-GitHubScriptsSelective {
	<#
	.SYNOPSIS
	Selectively syncs files from GitHub to local folder based on pattern matching

	.DESCRIPTION
	Downloads files from a GitHub repository folder with selective behavior:
	- Files matching AlwaysSyncPatterns: Always updated if content differs
	- Other files: Downloaded only if missing locally (preserves local modifications)

	Uses GitHub API with caching to minimize rate limit impact.
	Normalizes line endings to CRLF and uses ASCII encoding for text files.

	.PARAMETER LocalFolder
	Local directory to sync files to (required)

	.PARAMETER Owner
	GitHub repository owner (default: 'KnifMelti')

	.PARAMETER Repo
	GitHub repository name (default: 'SandboxStart')

	.PARAMETER GitHubPath
	Path within repository (default: 'Source/assets/scripts')

	.PARAMETER Branch
	Branch to sync from (default: 'master')

	.PARAMETER AlwaysSyncPatterns
	Array of wildcard patterns for files that should always be updated.
	Default: @('Std-*.ps1', 'Std-*.txt')

	.PARAMETER UseCache
	Use cached GitHub API responses (60-minute TTL)

	.EXAMPLE
	Sync-GitHubScriptsSelective -LocalFolder "C:\wsb" -UseCache

	Syncs files from GitHub to C:\wsb, always updating Std-*.ps1 files

	.EXAMPLE
	Sync-GitHubScriptsSelective -LocalFolder "C:\wsb" -AlwaysSyncPatterns @('Std-*.ps1', 'Template-*.txt') -Verbose

	Syncs with custom patterns and verbose logging

	.NOTES
	Silent failure on GitHub API errors (graceful degradation to local files)
	#>
	param(
		[Parameter(Mandatory=$true)]
		[string]$LocalFolder,

		[string]$Owner = 'KnifMelti',

		[string]$Repo = 'SandboxStart',

		[string]$GitHubPath = 'Source/assets/scripts',

		[string]$Branch = 'master',

		[string[]]$AlwaysSyncPatterns = @('Std-*.ps1', 'Std-*.txt'),

		[switch]$UseCache
	)

	# Ensure local folder exists
	if (!(Test-Path $LocalFolder)) {
		New-Item -Path $LocalFolder -ItemType Directory -Force | Out-Null
	}

	try {
		# Suppress progress bar
		$oldProgressPreference = $ProgressPreference
		$ProgressPreference = 'SilentlyContinue'

		# Get folder contents from GitHub API (all files, no filter)
		Write-Verbose "Syncing from GitHub: $Owner/$Repo/$GitHubPath"
		$files = Get-GitHubFolderContents `
			-Owner $Owner `
			-Repo $Repo `
			-Path $GitHubPath `
			-Branch $Branch `
			-UseCache:$UseCache

		if (!$files -or $files.Count -eq 0) {
			Write-Verbose "No files returned from GitHub (using cached/local files)"
			$ProgressPreference = $oldProgressPreference
			return
		}

		Write-Verbose "Found $($files.Count) files on GitHub"

		# Track statistics
		$downloadedCount = 0
		$updatedCount = 0
		$skippedCount = 0
		$unchangedCount = 0

		foreach ($file in $files | Where-Object { $_.type -eq 'file' }) {
			try {
				# Check if file matches AlwaysSyncPatterns
				$isAlwaysSync = $false
				foreach ($pattern in $AlwaysSyncPatterns) {
					if ($file.name -like $pattern) {
						$isAlwaysSync = $true
						break
					}
				}

				$localPath = Join-Path $LocalFolder $file.name

				if ($isAlwaysSync) {
					# Check for custom override header before syncing (applies to ALL AlwaysSyncPatterns files)
					if (Test-Path $localPath) {
						$firstLine = Get-Content -Path $localPath -TotalCount 1 -ErrorAction SilentlyContinue
						if ($firstLine -match '^\s*#\s*CUSTOM\s+OVERRIDE') {
							Write-Verbose "Skipping sync for custom override: $($file.name)"
							$skippedCount++
							continue
						}
					}

					# Check if Std-*.txt file was deleted by user (state=0 in .ini)
					if ($file.name -like 'Std-*.txt' -and -not (Test-Path $localPath)) {
						$config = Get-PackageListConfig -WorkingDir $WorkingDir
						$listName = [System.IO.Path]::GetFileNameWithoutExtension($file.name)
						if ($config.ContainsKey($listName) -and $config[$listName] -eq 0) {
							Write-Verbose "Skipping download for deleted package list: $($file.name)"
							$skippedCount++
							continue
						}
					}

					# Always-sync files: Download and compare, overwrite if different
					# Download from raw URL
					$remoteContent = (Invoke-WebRequest -Uri $file.download_url -UseBasicParsing).Content

					# Normalize remote content to CRLF for Windows
					$remoteContentNormalized = $remoteContent -replace "`r`n", "`n"  # First normalize to LF
					$remoteContentNormalized = $remoteContentNormalized -replace "`n", "`r`n"  # Then convert to CRLF

					if (Test-Path $localPath) {
						$localContent = Get-Content $localPath -Raw -ErrorAction SilentlyContinue

						# Compare normalized content
						if ($remoteContentNormalized -eq $localContent) {
							Write-Verbose "Unchanged: $($file.name) (already up to date)"
							$unchangedCount++
							continue
						}
						else {
							Write-Verbose "Updated: $($file.name) (content changed)"
							$updatedCount++
						}
					}
					else {
						Write-Verbose "Downloaded: $($file.name) (new file)"
						$downloadedCount++
					}

					# Save with CRLF line endings
					$remoteContentNormalized | Set-Content -Path $localPath -Encoding UTF8 -NoNewline -Force
				}
				else {
					# Other files: Only download if missing locally
					if (-not (Test-Path $localPath)) {
						# Download from raw URL
						$remoteContent = (Invoke-WebRequest -Uri $file.download_url -UseBasicParsing).Content

						# Normalize remote content to CRLF for Windows
						$remoteContentNormalized = $remoteContent -replace "`r`n", "`n"
						$remoteContentNormalized = $remoteContentNormalized -replace "`n", "`r`n"

						# Save with CRLF line endings
						$remoteContentNormalized | Set-Content -Path $localPath -Encoding UTF8 -NoNewline -Force

						Write-Verbose "Downloaded: $($file.name) (new file)"
						$downloadedCount++
					}
					else {
						Write-Verbose "Skipped: $($file.name) (preserving local file)"
						$skippedCount++
					}
				}
			}
			catch {
				Write-Verbose "Failed to sync $($file.name): $_"
				continue
			}
		}

		# Remote cleanup: Delete obsolete files
		# Only applies to Std-*.txt files (not Std-*.ps1 which are code)
		# and original default lists that have been replaced

		# 1. Cleanup obsolete Std-*.txt files
		$remoteStdTxtFiles = $files | Where-Object { $_.type -eq 'file' -and $_.name -like 'Std-*.txt' } | ForEach-Object { $_.name }
		$localStdTxtFiles = Get-ChildItem -Path $LocalFolder -Filter 'Std-*.txt' -File -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name

		if ($localStdTxtFiles) {
			$obsoleteStdFiles = $localStdTxtFiles | Where-Object { $_ -notin $remoteStdTxtFiles }

			foreach ($obsoleteFile in $obsoleteStdFiles) {
				$obsoletePath = Join-Path $LocalFolder $obsoleteFile
				$listName = [System.IO.Path]::GetFileNameWithoutExtension($obsoleteFile)

				# Check for CUSTOM OVERRIDE - preserve user modifications
				$firstLine = Get-Content -Path $obsoletePath -TotalCount 1 -ErrorAction SilentlyContinue
				if ($firstLine -match '^\s*#\s*CUSTOM\s+OVERRIDE') {
					Write-Verbose "Preserving custom override package list: $listName"
					continue
				}

				Write-Verbose "Removing obsolete Std-* package list: $listName"
				Remove-Item -Path $obsoletePath -Force -ErrorAction SilentlyContinue
				Set-PackageListConfig -ListName $listName -State 0 -WorkingDir (Split-Path $LocalFolder -Parent)
			}
		}

		# 2. Migration cleanup: Delete original default lists that have Std-* replacements
		$workingDir = Split-Path $LocalFolder -Parent
		$config = Get-PackageListConfig -WorkingDir $workingDir
		$migrationKeys = $config.Keys | Where-Object { $_ -like '_OriginalDefault_*' }

		foreach ($migrationKey in $migrationKeys) {
			$originalListName = $migrationKey -replace '^_OriginalDefault_', ''
			$originalPath = Join-Path $LocalFolder "$originalListName.txt"

			# Check if file still exists locally
			if (-not (Test-Path $originalPath)) {
				continue  # Already deleted
			}

			# Check for CUSTOM OVERRIDE (final protection)
			$firstLine = Get-Content -Path $originalPath -TotalCount 1 -ErrorAction SilentlyContinue
			if ($firstLine -match '^\s*#\s*CUSTOM\s+OVERRIDE') {
				Write-Verbose "Preserving custom override list: $originalListName"
				continue
			}

			# Check if Std-* replacement exists on GitHub
			$stdReplacementName = "Std-$originalListName.txt"
			if ($stdReplacementName -in $remoteStdTxtFiles) {
				Write-Verbose "Removing original default list (replaced by $stdReplacementName): $originalListName"
				Remove-Item -Path $originalPath -Force -ErrorAction SilentlyContinue
				Set-PackageListConfig -ListName $originalListName -State 0 -WorkingDir $workingDir

				# Remove migration tracking entry
				Set-PackageListConfig -ListName $migrationKey -State 0 -WorkingDir $workingDir
			}
		}

		# Summary
		Write-Verbose "Sync complete: $downloadedCount downloaded, $updatedCount updated, $unchangedCount unchanged, $skippedCount skipped"

		# Restore progress preference
		$ProgressPreference = $oldProgressPreference

	} catch {
		# Restore progress preference on error
		if ($oldProgressPreference) { $ProgressPreference = $oldProgressPreference }
		Write-Verbose "GitHub sync failed: $_"
		# Silent fail - fallback to local files
	}
}

#region Package List Management Functions

function Get-PackageListConfig {
	<#
	.SYNOPSIS
	Returns hashtable of package list states from .ini file

	.DESCRIPTION
	Reads the package-lists.ini file and returns a hashtable with list names and states.
	Creates default .ini file if missing.

	.OUTPUTS
	Hashtable with list names as keys and state (0 or 1) as values
	#>
	param(
		[string]$WorkingDir = $PWD.Path
	)

	$configPath = Join-Path (Join-Path $WorkingDir "wsb") "package-lists.ini"
	$config = @{}

	if (Test-Path $configPath) {
		try {
			$lines = Get-Content -Path $configPath -Encoding UTF8
			foreach ($line in $lines) {
				$line = $line.Trim()
				# Skip empty lines and comments
				if ([string]::IsNullOrWhiteSpace($line) -or $line.StartsWith('#')) {
					continue
				}

				# Parse key=value
				if ($line -match '^(.+?)=(.+)$') {
					$listName = $Matches[1].Trim()
					$state = $Matches[2].Trim()
					if ($state -match '^\d+$') {
						$config[$listName] = [int]$state
					}
				}
			}
		}
		catch {
			Write-Warning "Failed to read package list config: $_"
		}
	}

	return $config
}

function Set-PackageListConfig {
	<#
	.SYNOPSIS
	Updates .ini file with list name and state

	.PARAMETER ListName
	Package list name

	.PARAMETER State
	State: 1 = enabled, 0 = disabled/deleted

	.PARAMETER WorkingDir
	Working directory (defaults to current directory)
	#>
	param(
		[Parameter(Mandatory=$true)]
		[string]$ListName,

		[Parameter(Mandatory=$true)]
		[ValidateRange(0,1)]
		[int]$State,

		[string]$WorkingDir = $PWD.Path
	)

	$wsbDir = Join-Path $WorkingDir "wsb"
	if (-not (Test-Path $wsbDir)) {
		New-Item -ItemType Directory -Path $wsbDir -Force | Out-Null
	}

	$configPath = Join-Path $wsbDir "package-lists.ini"
	$config = Get-PackageListConfig -WorkingDir $WorkingDir

	# Update or add entry
	$config[$ListName] = $State

	# Write back to file
	try {
		$lines = @("# Package List Configuration")
		$lines += "# 1 = enabled, 0 = disabled/deleted"
		$lines += ""

		# Sort entries (but keep special entries at bottom)

		# Force arrays to prevent issues when there's only one key
		$specialKeys = @($config.Keys | Where-Object { $_ -like '_*' })
		$regularKeys = @($config.Keys | Where-Object { $_ -notlike '_*' } | Sort-Object)

		# Add regular keys first
		foreach ($key in $regularKeys) {
			$lines += "$key=$($config[$key])"
		}

		# Add special keys last
		foreach ($key in $specialKeys) {
			$lines += "$key=$($config[$key])"
		}
		Set-Content -Path $configPath -Value $lines -Encoding UTF8
	}
	catch {
		Write-Warning "Failed to write package list config: $_"
	}
}

function Initialize-PackageListConfig {
	<#
	.SYNOPSIS
	Scans wsb/ folder for .txt files, creates .ini entries

	.DESCRIPTION
	Creates initial package-lists.ini file if missing and populates it
	with all existing package lists (excluding script-mappings.txt).
	Called on GUI startup.

	.PARAMETER WorkingDir
	Working directory (defaults to current directory)
	#>
	param(
		[string]$WorkingDir = $PWD.Path
	)

	$wsbDir = Join-Path $WorkingDir "wsb"
	$configPath = Join-Path $wsbDir "package-lists.ini"

	# Only initialize if .ini doesn't exist
	if (Test-Path $configPath) {
		return
	}

	Write-Verbose "Initializing package list configuration"

	# Get all .txt files (excluding script-mappings.txt)
	$packageLists = @()
	if (Test-Path $wsbDir) {
		$txtFiles = Get-ChildItem -Path $wsbDir -Filter "*.txt" -File -ErrorAction SilentlyContinue
		foreach ($file in $txtFiles) {
			if ($file.Name -ne "script-mappings.txt") {
				$packageLists += $file.BaseName
			}
		}
	}

	# Create .ini with all lists enabled (state=1)
	foreach ($listName in $packageLists) {
		Set-PackageListConfig -ListName $listName -State 1 -WorkingDir $WorkingDir
	}
}

function Initialize-PackageListMigration {
	<#
	.SYNOPSIS
	One-time migration: Records current default list names for safe cleanup

	.DESCRIPTION
	Tracks which lists existed at time of migration. These can be safely
	deleted later when GitHub introduces Std-* replacements.
	Only runs once (tracked via _MigrationCompleted flag in .ini).

	.PARAMETER WorkingDir
	Working directory (defaults to current directory)
	#>
	param(
		[string]$WorkingDir = $PWD.Path
	)

	# Check if migration already completed
	$config = Get-PackageListConfig -WorkingDir $WorkingDir
	if ($config.ContainsKey('_MigrationCompleted')) {
		return  # Already migrated
	}

	Write-Verbose "Package list migration tracking initialized"

	# Define original default list names (shipped with SandboxStart)
	$originalDefaults = @('Python', 'AHK', 'AU3', '1. Missing in WSB')

	# Record which ones currently exist
	$wsbDir = Join-Path $WorkingDir "wsb"
	foreach ($listName in $originalDefaults) {
		$listPath = Join-Path $wsbDir "$listName.txt"
		if (Test-Path $listPath) {
			# Check if it has CUSTOM OVERRIDE
			$firstLine = Get-Content -Path $listPath -TotalCount 1 -ErrorAction SilentlyContinue
			if ($firstLine -notmatch '^\s*#\s*CUSTOM\s+OVERRIDE') {
				# Mark as original default (can be safely deleted during migration)
				Set-PackageListConfig -ListName "_OriginalDefault_$listName" -State 1 -WorkingDir $WorkingDir
				Write-Verbose "Tracked original default list: $listName"
			}
		}
	}

	# Mark migration as completed
	Set-PackageListConfig -ListName '_MigrationCompleted' -State 1 -WorkingDir $WorkingDir
}

#endregion

# Note: Functions are available via dot-sourcing (. script.ps1)
# No Export-ModuleMember needed when using dot-sourcing

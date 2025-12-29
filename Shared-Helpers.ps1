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

# Note: Functions are available via dot-sourcing (. script.ps1)
# No Export-ModuleMember needed when using dot-sourcing

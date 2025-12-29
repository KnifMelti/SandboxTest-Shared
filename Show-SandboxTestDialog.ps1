# User theme preference override (null = follow Windows, $true = force dark, $false = force light)
$script:UserThemeOverride = $null

function Get-ScriptMappings {
	<#
	.SYNOPSIS
	Reads script mapping configuration from external file

	.DESCRIPTION
	Loads script-to-pattern mappings from wsb\script-mappings.txt.
	Format: Pattern = ScriptName.ps1
	Example: InstallWSB.cmd = Std-WAU.ps1
	#>

	$mappingFile = Join-Path $Script:WorkingDir "wsb\script-mappings.txt"
	$mappings = @()

	# Create default mapping file if it doesn't exist
	if (-not (Test-Path $mappingFile)) {
		$wsbDir = Split-Path $mappingFile -Parent
		if (-not (Test-Path $wsbDir)) {
			New-Item -ItemType Directory -Path $wsbDir -Force | Out-Null
		}

		$defaultContent = @"
# Script Mapping Configuration for Windows Sandbox Testing
# Format: FilePattern = ScriptToExecute.ps1
#
# Patterns are evaluated in order. First match wins.
# Wildcards: * (any characters), ? (single character)
# The *.* pattern at the end acts as fallback.

InstallWSB.cmd = Std-WAU.ps1
*.installer.yaml = Std-Manifest.ps1
*.* = Std-Install.ps1
"@
		Set-Content -Path $mappingFile -Value $defaultContent -Encoding ASCII
	}

	# Read and parse mapping file
	try {
		$lines = Get-Content -Path $mappingFile -ErrorAction Stop

		# Migrate old script names to new ones
		$migrated = $false
		$updatedLines = @()

		foreach ($lineRaw in $lines) {
			$updatedLine = $lineRaw

			# Replace old script names with new ones
			if ($lineRaw -match '=\s*InstallWSB\.ps1\s*$') {
				$updatedLine = $lineRaw -replace 'InstallWSB\.ps1', 'Std-WAU.ps1'
				$migrated = $true
			}
			elseif ($lineRaw -match '=\s*WinGetManifest\.ps1\s*$') {
				$updatedLine = $lineRaw -replace 'WinGetManifest\.ps1', 'Std-Manifest.ps1'
				$migrated = $true
			}
			elseif ($lineRaw -match '=\s*Installer\.ps1\s*$') {
				$updatedLine = $lineRaw -replace 'Installer\.ps1', 'Std-Install.ps1'
				$migrated = $true
			}

			$updatedLines += $updatedLine
		}

		# Save migrated mappings back to file
		if ($migrated) {
			Set-Content -Path $mappingFile -Value ($updatedLines -join "`r`n") -Encoding ASCII

			# Delete old script files from wsb directory
			$wsbDir = Split-Path $mappingFile -Parent
			$oldScripts = @('InstallWSB.ps1', 'WinGetManifest.ps1', 'Installer.ps1')
			foreach ($oldScript in $oldScripts) {
				$oldPath = Join-Path $wsbDir $oldScript
				if (Test-Path $oldPath) {
					Remove-Item -Path $oldPath -Force -ErrorAction SilentlyContinue
				}
			}

			# Re-read the migrated lines
			$lines = $updatedLines
		}

		foreach ($line in $lines) {
			$line = $line.Trim()

			# Skip comments and empty lines
			if ($line.StartsWith('#') -or [string]::IsNullOrWhiteSpace($line)) {
				continue
			}

			# Parse: Pattern = Script.ps1
			if ($line -match '^\s*(.+?)\s*=\s*(.+?)\s*$') {
				$pattern = $matches[1].Trim()
				$script = $matches[2].Trim()

				# Validate script name ends with .ps1
				if ($script -like "*.ps1") {
					$mappings += @{
						Pattern = $pattern
						Script = $script
					}
				}
			}
		}
	}
	catch {
		Write-Warning "Failed to read script mappings: $($_.Exception.Message)"
	}

	# Ensure fallback exists
	if (-not ($mappings | Where-Object { $_.Pattern -eq "*.*" })) {
		$mappings += @{
			Pattern = "*.*"
			Script = "Std-Install.ps1"
		}
	}

	return $mappings
}

function Get-PackageLists {
	<#
	.SYNOPSIS
	Retrieves all package list files from wsb directory

	.DESCRIPTION
	Scans the wsb directory for .txt files (excluding script-mappings.txt)
	and returns their base names for use in the package list dropdown.

	.OUTPUTS
	Array of package list names (without .txt extension)
	#>

	$packageListDir = Join-Path $Script:WorkingDir "wsb"
	$lists = @()

	if (Test-Path $packageListDir) {
		$txtFiles = Get-ChildItem -Path $packageListDir -Filter "*.txt" -File -ErrorAction SilentlyContinue
		foreach ($file in $txtFiles) {
			# Exclude script-mappings.txt from package lists
			if ($file.Name -ne "script-mappings.txt") {
				$lists += $file.BaseName
			}
		}
	}

	return $lists | Sort-Object
}

function Get-PackageListTooltip {
	<#
	.SYNOPSIS
	Returns tooltip text based on whether package lists exist
	#>

	$lists = Get-PackageLists
	if ($lists.Count -eq 0) {
		return "No package lists found. Click '[Create new list...]' to create one."
	}
	return "Select a package list to install via WinGet"
}

function Show-PackageListEditor {
	<#
	.SYNOPSIS
	Shows dialog for creating or editing package lists or script mappings

	.PARAMETER ListName
	Optional list name to edit. If empty, creates new list. Not used in ScriptMapping mode.

	.PARAMETER EditorMode
	Editor mode: "PackageList" (default) or "ScriptMapping"

	.OUTPUTS
	Hashtable with DialogResult and ListName
	#>
	param(
		[string]$ListName = "",
		[ValidateSet("PackageList", "ScriptMapping")]
		[string]$EditorMode = "PackageList"
	)

	# Create editor form
	$editorForm = New-Object System.Windows.Forms.Form
	$editorForm.Text = switch ($EditorMode) {
		"PackageList" { if ($ListName) { "Edit Package List: $ListName" } else { "Create New Package List" } }
		"ScriptMapping" { "Edit Script Mappings" }
	}
	$editorForm.Size = switch ($EditorMode) {
		"PackageList" { New-Object System.Drawing.Size(420, 370) }
		"ScriptMapping" { New-Object System.Drawing.Size(600, 500) }
	}
	$editorForm.StartPosition = "CenterParent"
	$editorForm.FormBorderStyle = "FixedDialog"
	$editorForm.MaximizeBox = $false
	$editorForm.MinimizeBox = $false

	# Use the same icon as main form
	try {
		if ($Script:AppIcon) {
			$editorForm.Icon = $Script:AppIcon
			$editorForm.ShowIcon = $true
		} else {
			$editorForm.ShowIcon = $false
		}
	}
	catch {
		$editorForm.ShowIcon = $false
	}

	# Detect Windows theme preference (use same theme as main form)
	if ($null -ne $script:UserThemeOverride) {
		$useDarkMode = $script:UserThemeOverride
	}
	else {
		$useDarkMode = Get-WindowsThemeSetting
	}

	$y = 15
	$margin = 15
	$controlWidth = switch ($EditorMode) {
		"PackageList" { 380 }
		"ScriptMapping" { 560 }
	}

	# List name field - only for PackageList mode
	if ($EditorMode -eq "PackageList") {
		$lblListName = New-Object System.Windows.Forms.Label
		$lblListName.Location = New-Object System.Drawing.Point($margin, $y)
		$lblListName.Size = New-Object System.Drawing.Size(150, 20)
		$lblListName.Text = "List Name:"
		$editorForm.Controls.Add($lblListName)

		$txtListName = New-Object System.Windows.Forms.TextBox
		$txtListName.Location = New-Object System.Drawing.Point($margin, ($y + 20))
		$txtListName.Size = New-Object System.Drawing.Size($controlWidth, 23)
		$txtListName.Text = $ListName
		$txtListName.ReadOnly = ($ListName -ne "")
		$editorForm.Controls.Add($txtListName)

		$y += 50
	}

	# Content text area
	$lblPackages = New-Object System.Windows.Forms.Label
	$lblPackages.Location = New-Object System.Drawing.Point($margin, $y)
	$lblPackages.Size = New-Object System.Drawing.Size(400, 20)
	$lblPackages.Text = switch ($EditorMode) {
		"PackageList" { "Package IDs (one per line):" }
		"ScriptMapping" { "Script Mappings Configuration:" }
	}
	$editorForm.Controls.Add($lblPackages)

	$txtPackages = New-Object System.Windows.Forms.TextBox
	$txtPackages.Location = New-Object System.Drawing.Point($margin, ($y + 25))
	$txtPackages.Size = switch ($EditorMode) {
		"PackageList" { New-Object System.Drawing.Size($controlWidth, 140) }
		"ScriptMapping" { New-Object System.Drawing.Size($controlWidth, 270) }
	}
	$txtPackages.Multiline = $true
	$txtPackages.ScrollBars = "Vertical"
	$txtPackages.AcceptsReturn = $true
	$txtPackages.Font = New-Object System.Drawing.Font("Consolas", 9)
	$txtPackages.WordWrap = ($EditorMode -eq "PackageList")
	$editorForm.Controls.Add($txtPackages)

	# Load existing content if editing
	if ($EditorMode -eq "ScriptMapping") {
		$listPath = Join-Path (Join-Path $Script:WorkingDir "wsb") "script-mappings.txt"
	} else {
		$listPath = if ($ListName) { Join-Path (Join-Path $Script:WorkingDir "wsb") "$ListName.txt" } else { $null }
	}

	if ($listPath -and (Test-Path $listPath)) {
		try {
			$txtPackages.Text = (Get-Content -Path $listPath -Raw).Trim()
		}
		catch {
			[System.Windows.Forms.MessageBox]::Show("Error loading: $($_.Exception.Message)", "Load Error", "OK", "Error")
		}
	}

	$y += switch ($EditorMode) {
		"PackageList" { 175 }
		"ScriptMapping" { 320 }
	}

	# Help text
	$lblHelp = New-Object System.Windows.Forms.Label
	$lblHelp.Location = New-Object System.Drawing.Point($margin, $y)
	$lblHelp.Size = switch ($EditorMode) {
		"PackageList" { New-Object System.Drawing.Size($controlWidth, 50) }
		"ScriptMapping" { New-Object System.Drawing.Size($controlWidth, 70) }
	}
	$lblHelp.Text = switch ($EditorMode) {
		"PackageList" { "Example: Notepad++.Notepad++`nUse WinGet package IDs from winget search`nComments: Lines starting with # are ignored" }
		"ScriptMapping" { @"
Format: FilePattern = ScriptToExecute.ps1
Example: InstallWSB.cmd = InstallWSB.ps1
Patterns are matched against folder/file names (case-insensitive).
Wildcards: * (any characters), ? (single character)
Comments: Lines starting with # are ignored.
"@ }
	}
	$lblHelp.Name = 'lblHelp'  # For theme detection
	$editorForm.Controls.Add($lblHelp)

	$y += switch ($EditorMode) {
		"PackageList" { 50 }
		"ScriptMapping" { 90 }
	}

	# Buttons
	$btnSave = New-Object System.Windows.Forms.Button
	$btnSave.Location = New-Object System.Drawing.Point(($margin + $controlWidth - 160), $y)
	$btnSave.Size = New-Object System.Drawing.Size(75, 30)
	$btnSave.Text = "Save"
	$btnSave.Add_Click({
		if ($EditorMode -eq "PackageList") {
			# PackageList mode - validate and get list name
			$listNameValue = $txtListName.Text.Trim()

			# Validate list name
			if ([string]::IsNullOrWhiteSpace($listNameValue)) {
				[System.Windows.Forms.MessageBox]::Show("Please enter a list name.", "Validation Error", "OK", "Warning")
				return
			}

			# Check for invalid filename characters
			$invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
			if ($listNameValue.IndexOfAny($invalidChars) -ge 0) {
				[System.Windows.Forms.MessageBox]::Show("List name contains invalid characters.", "Validation Error", "OK", "Warning")
				return
			}

			# Prevent overwriting script-mappings.txt
			if ($listNameValue -eq "script-mappings") {
				[System.Windows.Forms.MessageBox]::Show("Cannot use reserved name 'script-mappings'.", "Validation Error", "OK", "Warning")
				return
			}

			$wsbDir = Join-Path $Script:WorkingDir "wsb"
			if (-not (Test-Path $wsbDir)) {
				New-Item -ItemType Directory -Path $wsbDir -Force | Out-Null
			}

			$listPath = Join-Path $wsbDir "$listNameValue.txt"
		} else {
			# ScriptMapping mode - direct to script-mappings.txt
			$wsbDir = Join-Path $Script:WorkingDir "wsb"
			if (-not (Test-Path $wsbDir)) {
				New-Item -ItemType Directory -Path $wsbDir -Force | Out-Null
			}
			$listPath = Join-Path $wsbDir "script-mappings.txt"
			$listNameValue = "script-mappings"
		}

		# Save the file
		try {
			$packageContent = $txtPackages.Text.Trim()
			Set-Content -Path $listPath -Value $packageContent -Encoding ASCII

			$script:__editorReturn = @{
				DialogResult = 'OK'
				ListName = $listNameValue
			}
			$editorForm.Close()
		}
		catch {
			[System.Windows.Forms.MessageBox]::Show("Error saving: $($_.Exception.Message)", "Save Error", "OK", "Error")
		}
	})
	$editorForm.Controls.Add($btnSave)

	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Location = New-Object System.Drawing.Point(($margin + $controlWidth - 75), $y)
	$btnCancel.Size = New-Object System.Drawing.Size(75, 30)
	$btnCancel.Text = "Cancel"
	$btnCancel.Add_Click({
		$script:__editorReturn = @{ DialogResult = 'Cancel' }
		$editorForm.Close()
	})
	$editorForm.Controls.Add($btnCancel)

	$editorForm.AcceptButton = $btnSave
	$editorForm.CancelButton = $btnCancel

	# Apply theme based on Windows settings
	if ($useDarkMode) {
		Set-DarkModeTheme -Control $editorForm
		Set-DarkTitleBar -Form $editorForm -UseDarkMode $true
	}
	else {
		Set-LightModeTheme -Control $editorForm
		Set-DarkTitleBar -Form $editorForm -UseDarkMode $false
	}

	[void]$editorForm.ShowDialog()

	if ($script:__editorReturn) {
		return $script:__editorReturn
	} else {
		return @{ DialogResult = 'Cancel' }
	}
}

# Determine the appropriate script based on selected file or directory contents
function Find-MatchingScript {
	param(
		[string]$Path,
		[string]$FileName = $null
	)
	
	$mappings = Get-ScriptMappings
	
	# If specific file selected, test against patterns
	if ($FileName) {
		foreach ($mapping in $mappings) {
			if ($FileName -like $mapping.Pattern) {
				return $mapping.Script
			}
		}
	}
	
	# If no file or no match, scan directory for pattern matches
	if (Test-Path $Path) {
		# Exclude *.* fallback from directory scan
		$scanMappings = $mappings | Where-Object { $_.Pattern -ne "*.*" }
		
		foreach ($mapping in $scanMappings) {
			$matchingFiles = Get-ChildItem -Path $Path -Filter $mapping.Pattern -File -ErrorAction SilentlyContinue
			if ($matchingFiles) {
				return $mapping.Script
			}
		}
	}
	
	# Fallback to last mapping (should be *.*)
	$fallback = $mappings | Where-Object { $_.Pattern -eq "*.*" } | Select-Object -First 1
	if ($fallback) {
		return $fallback.Script
	} else {
		return "Std-Install.ps1"
	}
}

# Helper function to fetch stable WinGet versions from GitHub
function Get-StableWinGetVersions {
	<#
	.SYNOPSIS
	Fetches the 25 most recent stable WinGet versions from GitHub
	
	.DESCRIPTION
	Queries the GitHub API for microsoft/winget-cli releases and returns
	the tag names of the 25 most recent stable (non-prerelease) versions
	that have assets available. Excludes releases without assets to prevent
	installation failures.
	
	.OUTPUTS
	Array of version strings (e.g., "v1.7.10514", "v1.7.10582")
	#>
	try {
		# Request 100 releases to ensure we get 25 stable ones after filtering pre-releases and checking assets
		# Use GitHub API helper with caching and fallback
		Write-Verbose "Fetching WinGet releases from GitHub API..."

		$releases = Get-GitHubReleases `
			-Owner "microsoft" `
			-Repo "winget-cli" `
			-PerPage 100 `
			-StableOnly `
			-UseCache

		# Filter to only releases that have assets and get top 25
		$stableReleases = $releases | Where-Object {
			($_.assets) -and
			($_.assets.Count -gt 0)
		} | Select-Object -First 25

		# Extract tag names (e.g., "v1.7.10514")
		$versions = $stableReleases | ForEach-Object { $_.tag_name }

		Write-Verbose "Found $($versions.Count) stable WinGet versions with assets"
		return $versions
	}
	catch {
		Write-Warning "Failed to fetch WinGet versions from GitHub: $($_.Exception.Message)"
		return @()
	}
}

# ============================================================================
# Dark Mode Theme Functions
# ============================================================================
# Note: Get-WindowsThemeSetting is in Shared-Helpers.ps1 (shared with SandboxTest.ps1)

function Set-DarkModeTheme {
	<#
	.SYNOPSIS
	Applies dark theme to a Windows Form and all its controls recursively

	.PARAMETER Control
	The form or control to apply dark theme to

	.PARAMETER UpdateButtonBackColor
	Optional. BackColor override for the update button (adaptive green)
	#>

	param(
		[Parameter(Mandatory)]
		[System.Windows.Forms.Control]$Control,

		[Parameter(Mandatory=$false)]
		[System.Drawing.Color]$UpdateButtonBackColor = [System.Drawing.Color]::Empty
	)

	# Dark mode color palette
	$darkBg = [System.Drawing.Color]::FromArgb(32, 32, 32)
	$darkFg = [System.Drawing.Color]::White
	$darkButtonBg = [System.Drawing.Color]::FromArgb(70, 70, 70)
	$darkTextBoxBg = [System.Drawing.Color]::FromArgb(45, 45, 45)
	$darkGrayText = [System.Drawing.Color]::FromArgb(180, 180, 180)  # Lighter gray for dark mode

	# Apply base colors
	$Control.BackColor = $darkBg
	$Control.ForeColor = $darkFg

	# Special handling by control type
	if ($Control -is [System.Windows.Forms.Button]) {
		$Control.BackColor = $darkButtonBg

		# Special case: Update button with adaptive green
		if ($Control.Name -eq 'btnUpdate' -or ($Control.Text -eq [char]0x2B06)) {
			if ($UpdateButtonBackColor -ne [System.Drawing.Color]::Empty) {
				$Control.BackColor = $UpdateButtonBackColor
			}
			else {
				# Dark mode adaptive green (darker, more muted)
				$Control.BackColor = [System.Drawing.Color]::FromArgb(60, 120, 60)
			}
		}
	}
	elseif ($Control -is [System.Windows.Forms.TextBox]) {
		$Control.BackColor = $darkTextBoxBg
		$Control.ForeColor = $darkFg
		# Preserve font settings (Consolas for script editor)
	}
	elseif ($Control -is [System.Windows.Forms.ComboBox]) {
		$Control.BackColor = $darkTextBoxBg
		$Control.ForeColor = $darkFg
	}
	elseif ($Control -is [System.Windows.Forms.CheckBox]) {
		# CheckBox needs parent background color
		$Control.BackColor = $darkBg
		$Control.ForeColor = $darkFg
	}
	elseif ($Control -is [System.Windows.Forms.Label]) {
		# Check if this is the help label (gray text)
		if ($Control.Name -eq 'lblHelp' -or $Control.ForeColor.ToArgb() -eq [System.Drawing.Color]::Gray.ToArgb()) {
			$Control.ForeColor = $darkGrayText
		}
		else {
			$Control.ForeColor = $darkFg
		}
	}

	# Recursively apply to child controls
	foreach ($child in $Control.Controls) {
		Set-DarkModeTheme -Control $child -UpdateButtonBackColor $UpdateButtonBackColor
	}
}

function Set-LightModeTheme {
	<#
	.SYNOPSIS
	Applies light theme to a Windows Form and all its controls recursively

	.PARAMETER Control
	The form or control to apply light theme to

	.PARAMETER UpdateButtonBackColor
	Optional. BackColor override for the update button (adaptive green)
	#>

	param(
		[Parameter(Mandatory)]
		[System.Windows.Forms.Control]$Control,

		[Parameter(Mandatory=$false)]
		[System.Drawing.Color]$UpdateButtonBackColor = [System.Drawing.Color]::Empty
	)

	# Light mode color palette
	$lightBg = [System.Drawing.Color]::FromArgb(240, 240, 240)
	$lightFg = [System.Drawing.Color]::Black
	$lightButtonBg = [System.Drawing.Color]::LightGray
	$lightTextBoxBg = [System.Drawing.Color]::White
	$lightGrayText = [System.Drawing.Color]::Gray

	# Apply base colors
	$Control.BackColor = $lightBg
	$Control.ForeColor = $lightFg

	# Special handling by control type
	if ($Control -is [System.Windows.Forms.Button]) {
		$Control.BackColor = $lightButtonBg

		# Special case: Update button with adaptive green
		if ($Control.Name -eq 'btnUpdate' -or ($Control.Text -eq [char]0x2B06)) {
			if ($UpdateButtonBackColor -ne [System.Drawing.Color]::Empty) {
				$Control.BackColor = $UpdateButtonBackColor
			}
			else {
				# Light mode adaptive green (brighter)
				$Control.BackColor = [System.Drawing.Color]::LightGreen
			}
		}
	}
	elseif ($Control -is [System.Windows.Forms.TextBox]) {
		$Control.BackColor = $lightTextBoxBg
		$Control.ForeColor = $lightFg
		# Preserve font settings (Consolas for script editor)
	}
	elseif ($Control -is [System.Windows.Forms.ComboBox]) {
		$Control.BackColor = $lightTextBoxBg
		$Control.ForeColor = $lightFg
	}
	elseif ($Control -is [System.Windows.Forms.CheckBox]) {
		# CheckBox needs parent background color
		$Control.BackColor = $lightBg
		$Control.ForeColor = $lightFg
	}
	elseif ($Control -is [System.Windows.Forms.Label]) {
		# Check if this is the help label (gray text)
		if ($Control.Name -eq 'lblHelp' -or $Control.ForeColor.ToArgb() -eq [System.Drawing.Color]::Gray.ToArgb()) {
			$Control.ForeColor = $lightGrayText
		}
		else {
			$Control.ForeColor = $lightFg
		}
	}

	# Recursively apply to child controls
	foreach ($child in $Control.Controls) {
		Set-LightModeTheme -Control $child -UpdateButtonBackColor $UpdateButtonBackColor
	}
}

function Set-DarkTitleBar {
	<#
	.SYNOPSIS
	Sets the window title bar to dark or light mode using DwmSetWindowAttribute

	.PARAMETER Form
	The Windows Form to apply title bar theming to

	.PARAMETER UseDarkMode
	If $true, use dark title bar. If $false, use light title bar.
	#>

	param(
		[Parameter(Mandatory)]
		[System.Windows.Forms.Form]$Form,

		[Parameter(Mandatory)]
		[bool]$UseDarkMode
	)

	# Define DwmSetWindowAttribute if not already defined
	if (-not ([System.Management.Automation.PSTypeName]'DarkMode.DwmApi').Type) {
		Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
namespace DarkMode {
	public class DwmApi {
		[DllImport("dwmapi.dll")]
		public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
	}
}
"@ -ErrorAction SilentlyContinue
	}

	# Capture parameter value for use in scriptblock
	$darkMode = $UseDarkMode

	# Apply dark title bar via Shown event (window handle must be created first)
	$Form.Add_Shown({
		try {
			# DWMWA_USE_IMMERSIVE_DARK_MODE = 20
			$titleBarMode = if ($darkMode) { 1 } else { 0 }
			[DarkMode.DwmApi]::DwmSetWindowAttribute($this.Handle, 20, [ref]$titleBarMode, 4) | Out-Null
		}
		catch {
			# Silent fail - not critical if title bar theming doesn't work
			Write-Verbose "Failed to set title bar theme: $($_.Exception.Message)"
		}
	}.GetNewClosure())
}

# Global theme toggle handler - must be in script scope to be accessible from event handlers
$global:ToggleFormThemeHandler = {
	param($FormControl, $UpdateButtonColor)

	# Toggle the override state
	if ($null -eq $script:UserThemeOverride) {
		# First toggle: invert current Windows theme
		$windowsPrefersDark = Get-WindowsThemeSetting
		$script:UserThemeOverride = -not $windowsPrefersDark
	}
	else {
		# Subsequent toggles: flip the override
		$script:UserThemeOverride = -not $script:UserThemeOverride
	}

	# Apply the new theme
	if ($script:UserThemeOverride) {
		Set-DarkModeTheme -Control $FormControl -UpdateButtonBackColor $UpdateButtonColor
		Set-DarkTitleBar -Form $FormControl -UseDarkMode $true
	}
	else {
		Set-LightModeTheme -Control $FormControl -UpdateButtonBackColor $UpdateButtonColor
		Set-DarkTitleBar -Form $FormControl -UseDarkMode $false
	}

	# Refresh the form
	$FormControl.Refresh()
}

# Define the dialog function here since it's needed before the main functions section
function Show-SandboxTestDialog {
	<#
	.SYNOPSIS
	Shows a GUI dialog for configuring Windows Sandbox test parameters

	.DESCRIPTION
	Creates a Windows Forms dialog to collect all parameters needed for SandboxTest function
	#>

	# Embedded icon data (Base64-encoded from Source\assets\icon.ico)
	# Generated: 2025-12-20
	# Original size: 15KB (.ico) ??? 20KB (base64)
	$iconBase64 = @"
AAABAAMAMDAAAAEAIACoJQAANgAAACAgAAABACAAqBAAAN4lAAAQEAAAAQAgAGgEAACGNgAAKAAAADAAAABgAAAAAQAgAAAAAACAJQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA/78/BP+/PwQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP/MMw/+yzNP/sgun/7EKp3+wCVK/rAnDQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD+2jYO/tE6T/7RN7P+zTPy/8gt///CKP//viLw/rserv63G0v+sBMNAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA/to2Dv7ROk/+0jiz/tM58f/ROP//zjP//8gt///CJ///vSH//7kd//+3G/D+uBqu/rYbSv6wEw0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP7aNg7+0TpP/tI6s/7TOvH/0jn//9I4///RN///zTL//8cs///BJv//vCD//7gc//+3G//+thr+/7UZ8P60GK/+thhK/rATDQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD+2jYO/tE6T/7UOrP+0zrx/9M6///SOf//0jj//9I4///RN///zTL//8cs///BJv//vCD//7gb//+2Gv//thn//7UY//+0F///sxfw/rMVr/6zFEr+sBMNAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA/to2Dv7ROk/+1Dqz/tQ68f/TOv//0zn//9I5///SOP//0jj//9I4///RN///zTP//8gt///CJ///vSH//7gc//+2Gv//tRn//7QY//+0F///sxb//rIV/v+xFPD+sRSu/rMRSv6wEw0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP/MMw/+1TpP/tQ6s/7TO/L+0zr+/9M5///TOf//0zn//9I5///SOf//0jj//9I5///SOP//zjP//8gt///DKP//vSL//7kd//+3G///thn//7UY//+zF///shX//7EU//+wE///sBL//68S8P6uEa/+rxFK/rATDQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD+2jYO/tU6T/7UO7P+0zvy/9M7///TOv//0zr//9M5///TOf//0zn//9I5///TOf//0zn//9M5///SOP//zjT//8ku///EKf//vyP//7oe//+4HP//txr//7UZ//+0F///shX//7EU//+wE///rxL//64R//6tEP7/rQ/w/qwOrv6sDUr+sBMNAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA/to2Dv7VPU/+1Duz/tM78v/UO///0zr//9M6///TOv//0zr//9M6///TOf//0zr//9M6///TOv//0zr//9Q6///TOv//0DX//8sw///GK///wSX//7wg//+6Hv//uBz//7Ya//+1GP//sxb//7IU//+wE///rxL//64Q//+tD///rA7//6sO//+rDfD+qg2v/qwNSv6cEw0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP/MRA/+1T1P/tU7s/7UPPL/1Dz//9Q7///TOv//0zr//9M6///TOv//0zr//9M6///UOv//0zr//9Q6///UO///1Dv//9U7///UO///0Tf//8wy///ILf//wyj//74j//+8If//uh7//7gc//+2Gv//tBj//7MW//+xE///rxL//64R//+tD///rA7//6sN//+qDP//qQv//6gK8P6nCq/+qApK/rATDQAAAAAAAAAAAAAAAAAAAAD+2kgO/tU9T/7VPbP+1Dzy/9Q8///UO///1Dv//9Q7///UO///0zr//9M6///TOv//0zv//9Q6///UO///1Dv//9Q7//7VPP/+1Tz//tU8///VPP//0jn//840//7KL//+xSr//8Em//6/JP//vCH//7of//+4HP//thr//7QX//+yFf//sBP//64R//+tD///rA7//6sN//+qC///qQr//6gJ//+nCf//pwjw/qcIr/6mCkv+sBMNAAAAAP/XP0D+1z6z/tU+8v/VPf//1Tz//9Q7///UO///1Dv//9Q7///UOv//1Dv//9Q7//7UO//+1Dv//tQ7///UPP/+1Tz//9U8//7WPf/+1j3//tY+///WPv//0zr//c41//7KMf/+xy3//sMp///BJv//vyT//70i//+6H//+uBz//rYa//+0F///shX//7AT//+uEf//rQ///6sN//+qDP//qAr//6cJ//+nCP//pgj//6cI//+oCfD+qQuu/qwMPv7XQNv/1j7//9Y+///VPf//1Dz//9Q7///UO///1Dv//9Q7///UO///1Dv//9Q7///UPP/+1Tz//9U8//7WPf//1j3//9Y+//7WPv/+1z///tc///vSPP/uuyv/3Z8Y/92dF//vsyL//MIp///EKf//wif//sAl//69Iv//uyD//rkd//+2Gv//tBj//rIV//+wEv//rhH//6wO//+qDP//qQr//6gJ//+nCP//pgf//6YI//+nCf/+qQv+/qsO2v/YQP//1z///9Y9///VPf//1Tz//9Q7///UO///1Dz//9U8///VPP//1Tz//9U8///VPf//1j3//tY+//7XPv//1z7//9c////YP//80z3/774v/9qcGP/Ohgn/zIUM/8uNIP/PixX/25gU/++xIP/7wCj//sIn//7AJf//viP//7wg//65Hf//txv//rUY//+yFf//sBP//64Q//+sDv//qgz//6gK//+nCP//pgf//6YH//+nCf//qQv//qwO/v/YQP//1z///9Y+///VPf//1T3//9U8///VPP//1Tz//9U9///VPf//1j3//9Y+///WPv//1z7//9c////YP///2ED//NQ+/+/JSP/Jwnj/vqtn/86NHP/QiA//vqhj/3+9wP+KuK3/taZo/8+IDf/blxL/77Af//y/Jv//wCb//74j//68If//uh7//7gb//+1Gf//sxb//7AT//+uEf//rA7//6oM//+oCv//pwj//6cI//+oCf//qQv//6wO///YQP//1z///9Y+///VPf//1T3//9U8///VPf//1j3//9Y+///XPv//1z7//9c///7XP///2ED//9hA//zUPf/wvzD/3Ko2/5rIuP9dxN//WbzY/4i2rv+btZ7/dMbV/02+3/9Lutv/c7e+/8uSKv/Ogwb/z4YI/9yWEf/wrx7//L4l//6/JP//vSH//7sf//+4HP//thr//7MX//+xE///rhH//6wO//+qDP//qQr//6gJ//+pCv//qgz//6wP///YQP//1z///9Y+///WPv//1j7//9Y9///WPv//1z7//9c+///XP///2D///9hA//7YQP/81D7/8cEx/9+hHf/UjQ//wapi/2fT6f9Tz+z/Sr/h/0m32/9Lv+D/V9z1/1jd9v9Rzur/TbXT/5q0nP/FoE3/w5xG/8mWM//RiQ3/3JgT//CwHv/8vCP//r0i//+7IP//uR3//7ca//+0F///sRT//68R//+sD///qw3//6oM//+qDP//qw3//60Q///YQP//2D///9c////XP///1z7//9c////XP///2D///9hA///YQP//2UH//dU+//HDM//hpSH/15IT/9WOEf/Wkx3/r8Wb/1/n/P9b5v3/V973/1XW8f9Y4Pf/WOj+/1jo/v9V4vv/RL3e/0m11f9ct8//WrjR/3m7wP/IoUv/0YkM/9ONDv/emxX/8bAe//y7Iv//vCD//7oe//+3G///tRj//7IV//+wEv//rhD//60P//+sDv//rQ///68R///ZQf//2ED//9hA///YQP//10D//9hA///YQP//2UD//9lB//rSPv/wwjT/46kk/9qYGP/ZlBX/2ZMU/9GnSv+TwbX/ctDe/1nm/f9Z5/7/WOj+/1jo/v9Y6P7/Wej+/1bm//9W5v//VNn0/0zD4v9Iv+D/SsDf/0m00v+Wt6T/1pIa/9SPD//UjxD/1ZIS/+CfGP/urR3/+bcf//+6Hv//uBz//7YZ//+zFv//sRT//7AS//+vEv//rxL//7AT///aQv//2UH//9hB///YQf//2UH//9lB///ZQf/60j7/5LMv/8GBGP+1bw7/yIQT/9mVF//dmBj/3ZkZ/7K+kv9UyOj/VNPw/1nn/f9Y6P7/Wej+/2jn9/964eT/feLm/2zm9v9a5v7/V+X+/1fj/P9V4/z/Vd/5/1HD4P+otY//2ZUW/9mTEv/YlBP/048S/716Dv+gXQn/rWkM/9yYGP/5tR7//7kd//+2Gv//tRj//7MW//+yFf//shX//7IW///aQv//2kL//9lC///ZQf//2UH/+tM+/+a0L//Fgxf/sGQI/6pdBv+mWwf/pFsI/7BqDP/JhRT/26Au/5Hb0v9a4vr/Wef9/1jo/v9Y5/3/h97g/8m5cf/XoTb/2KE2/8+vWf+qz67/auT0/1Xl/v9V5P7/VOP+/1/X7f/FsGX/3JcV/9aSFP/AfA//nVoI/4lFBP+HQQL/iUIC/5FJBP+uaAv/3ZcW//mzHP//uBv//7Ya//+1Gf//tRj//7UZ///aQ///2kP//9pC//vTP//ntS//yIUW/7VnB/+wYAX/rl8F/61eBv+qXAf/plsH/6FYB/+gWQj/rHgq/3jc4P9Z6P7/WOj+/1fo/v9U2vT/pMGt/96aH//dlBP/3JMS/9yTEv/cmB7/vsKG/2bk9/9V4/7/VOD8/0a62P+CuLb/vqVg/59fD/+IRAT/hkED/4hCAv+KQwL/i0MC/4xDAv+MQwH/kkkD/69oC//dlxb/+bMc//+4HP//txv//7cb///bQ//71D//6bYv/8uHFv+5aQb/tGIE/7NiBP+xYQT/sGAF/65fBv+sXgb/ql0H/6ZbCP+iWQf/oWEZ/5Wznv9r5PT/WOj+/1jo/v9Nz+z/a7fJ/9GqT//fmBf/3pcV/96XFf/elxX/3p8n/5nYxv9V4/7/U+D8/0W82/9Gr8//bK27/45YJf+IQgP/ikMD/4xDAv+MRAL/jUQC/41DAf+NQwH/jUMB/4xCAf+SSQP/r2gK/92XFv/5sxz//7kd/+6+M//PiRX/vWsG/7hlA/+3ZAP/tmMD/7VjBP+zYgX/sWEF/7BgBv+uXwb/q10H/6hcB/+lWgj/olkI/6FfFf+ip4j/Yub7/1fm/v9P1vP/RLLY/3u5vv/LsGP/4KEp/+CbGv/gnBv/3aMx/6LPt/9W4/7/VOP+/1Tb9/9PxOL/baWv/49SGf+LQwL/jEMC/41EAv+NRAL/jkQC/45EAv+ORAH/jkMB/45DAf+NQwH/jUMB/5NJA/+waAr/450X/8d4Cv69ZwH/u2YC/7pmAv+5ZQP/uGQD/7ZjBP+0YgX/smEF/7BgBv+uXwb/rF4H/6lcB/+nWwj/pFoI/6JZCv+fm3j/YOP6/1bm//9U4vz/SMTm/0Kz2v9atdH/iLm1/6Kzj/+itpT/hcDB/17T7P9U4/7/VeP+/1Ti/v9U3Pj/g5yT/41HCP+MQwL/jEMC/41EAv+ORAL/j0QC/49EAf+PRAH/j0QB/49EAf+PQwH/jkMB/45DAf+OQwH/nlQF/75oANq9ZwH/vGcC/7tmAv+6ZQP/uWQD/7djBP+1YgT/s2EF/7FgBv+vXwb/rF4H/6pdB/+nWwj/pVoI/6diFv+Rx7//WOb//1bm//9V5f7/VN/6/0vH6f9Ct97/QLLa/0Oy2v9Fud//T8vs/1Xg+/9U4/7/VOH9/1zh+v9p1uj/knxY/4xEAv+MRAL/jUQC/45EAv+PRAL/j0QC/49EAv+PRAH/kEQB/5BEAf+PRAH/j0QB/49DAf+PQwH/jkMA28FmAD6+aAGuvWcC8LxnAv+7ZgP/uWUD/7dkBP+2YwT/s2IF/7FgBv+vXwb/rV4H/6tdB/+oXAj/plsI/6VcDf+gn3n/ZuT3/1fl/v9Y5f7/VeX+/1Ti/f9S2vf/T9Lx/1DS8f9T2fb/VeL8/1Tj/v9V4v7/Tsvo/4yysv+TbT//j0sN/41EAv+NRAL/jUQC/49EAv+PRAL/j0UC/5BEAf+QRAH/kEQB/5BEAf+QRAH/kEQB/o9DAfGPQgGzj0MAQAAAAADEYgANumUDS71mAa67ZgPwuWUD/rhkBP+2YwX/tGIF/7FhBv+wYAb/rl8H/6tdB/+pXAj/p1sI/6VaCf+mYhf/n598/5OzoP+Zuaj/eNvn/1bj/v9V4/7/VuP+/1Tj/v9U4/7/VeP+/1Pi/v9T4f3/T7vY/5KJc/+ORgX/jUQD/41EAv+ORAL/jkQC/49EAv+PRAL/kEUC/5BFAv+QRAH/kEQB/5BEAf+QQwHyj0QBs45DAE+RSAAOAAAAAAAAAAAAAAAAAAAAALBiAA29ZwNKumUCrrhkBO+2YwX/tGIF/7JhBf+wYAb/rl8H/6xeB/+qXQj/qVwJ/6ZbCf+kWgn/o1kL/6FZDP+hWhD/qY9i/23g8f9U4/7/VOP+/1zi+v9t2ej/aN7y/1bi/v9U4v3/cc7g/5lrPP+ORQP/jkUC/45FAv+PRQL/j0UC/5BFAv+QRQL/kEUC/5BFAf+QRAH/kEQB8o9EAbOOQwBPiEQADwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAsGIADbpjA0q3YwSutGEF8LNhBv+xYAb/r18H/61eB/+rXQj/qVwI/6dbCf+lWgn/pFkK/6JYC/+gVwv/omAa/5C9sf9n3vH/YeL5/46ypv+bazf/nYFZ/4q8t/+DuLL/mX9X/5JMDf+ORQP/j0UC/49FAv+QRQL/kEUC/5BFAv+QRQL/kEUC/5BEAvKPRAGzkUMAT5FIAA4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACwYgANtGIGS7RgBa6xXwbwr18H/65fB/+sXgj+ql0I/6hcCf+mWwr/pVoK/6NZC/+hWAv/oFgM/6NnJf+ffEn/no1n/5tfIv+TSgX/kEcF/5NPEP+TTg7/j0YD/49GA/+PRgP/kEYC/5BGAv+QRgL/kEUC/5BFAv+QRALyj0QBs45DAE+RSAAOAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALBiAA2zYAZKsWAHrq5eB/CtXgj+q14J/6lcCf+oWwn/ploK/6RZCv+jWQv/oVgM/59XDP+dVAv/mlEJ/5dOB/+USgX/kUgE/5FHA/+RRgP/kEYD/5FGA/+RRgP/kUYD/5BGAv+RRgL/kEUC8pFEAbOTRANOiEQADwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAsGIADa9gBkqtXgevrF0I8KpdCf+pXAn/p1sK/6VaC/+kWQv/olkM/6FXDP+eVQv/m1EJ/5hOB/+VSwX/kkgE/5FHA/+SRwP/kUcD/5FHA/+RRwP/kUYD/5BFAvKRRQGzkUMDT5FIAA4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACwYgANqVsKS6tdCK6oXAnwp1sK/6ZbC/+lWgv/o1kM/6FYDP+fVQv/nFIJ/5lOB/+WSwX/k0kE/5JIA/+SRwP/kkcD/5JHA/+RRgPykUUCs5FHA0+RSAAOAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALBiAA2oXQpKp1sKr6ZbCvCmWgv+pFkM/6JYDP+gVgv/nVIJ/5pPB/+XSwX/lEkE/5NIA/+TSAP+k0YD8ZJHArORRwNPiEQADwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAnGIADahZCkqnWguupFkL8KNZDP6gVgv/nVIJ/5pPB/+XTAX/lEkE/5RHA/GSSAKzkUcDT5FIAA4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACcYhMNpVkKSqRZC66gVgrwnlMJ/5tPB/+XSwXylEoEs5RHA0+RSAAOAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJxOEw2hVgpKnlIJnptQCJ+ZTgZOkUgADgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAvz8ABL8/AAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP///////wAA////////AAD///////8AAP///////wAA///+f///AAD///gf//8AAP//4Af//wAA//+AAf//AAD//gAAf/8AAP/4AAAf/wAA/+AAAAf/AAD/gAAAAf8AAP4AAAAAfwAA+AAAAAAfAADgAAAAAAcAAIAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACAAAAAAAEAAOAAAAAABwAA+AAAAAAfAAD+AAAAAH8AAP+AAAAB/wAA/+AAAAf/AAD/+AAAH/8AAP/+AAB//wAA//+AAf//AAD//+AH//8AAP//+B///wAA///+f///AAD///////8AAP///////wAA////////AAD///////8AACgAAAAgAAAAQAAAAAEAIAAAAAAAgBAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//8AAf7RNhz/yS9g/sMoX/68JRv//wABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//8AAf7RNhz+0zlv/tA2z/7ILvr+wSX5/7sfzP64HGz+sxwb//8AAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//8AAf7aPxz+0zlv/tM4zv7SOfr/0Db//8ku///AJf//uR3//rca+f+2Gsz+tRds/rMSG///AAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//8AAf7aPxz+0zlv/tI5z/7TOfr/0jn//9I4///QNv//yC7//8Ak//+5Hf//thr//7UZ//60F/n/shbM/rEVbP6zEhv//wABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//8AAf7aPxz+1Ttv/tM5z/7TOvr/0zn//9I5///SOP//0jj//9E3///JL///wSb//7oe//+3Gv//tRj//7MW//+yFf/+sRP5/7ASzP6uEGz+qRIb//8AAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//8AAf7aPxz+1Ttv/tM7z/7TO/r/0zr//9M5///TOf//0zn//9M5///TOf//0Tj//8sx///DKP//vCD//7gc//+2Gf//tBf//7IU//+wE///rxH//q4P+f+sD8z+rA5s/qkJG///AAEAAAAAAAAAAAAAAAAAAAAA//8AAf7aPxz+1Ttv/tM8z/7UO/r/1Dr//9M6///TOv//0zr//9M6///TOv//1Dr//9Q7///TOf//zTP//8Yr//+/I///ux///7gc//+1Gf//sxb//7AT//+uEf//rQ///6sN//6rDPn/qgvM/qcJbP6pCRv//wABAAAAAP7ZQhv+1T5v/tU8z/7UPPr/1Dv//9Q7///UOv//1Dv//9Q7///UO///1Dv//9Q8///VPP//1T3//9U8///PNv//yS///sIo//+/JP//vCD//7gc//+1Gf//shX//68S//+tD///qw3//6kL//+oCf/+pwj5/6cIzP6nCWz+rQoZ/tc+w/7WPvr/1T3//9Q7///UO///1Dv//9Q7///UO///1Dz//9U8//7VPf/+1j3//tY+//7XPv/5zTf/67Ml/+uvIf/5vif//sIo///AJP//vCH//7kd//+1Gf//shX//68S//+sD///qgz//6gJ//+mCP//pgj//qgK+f6rDcH+1z/+/9Y+///VPf//1Dz//9Q8///VPP//1Tz//9U9///WPf//1j7//9c+//7XP//50kL/5r0//9aUFP/LjBz/uptJ/82cMf/oqR3/+bwl//7AJf//vSL//7ke//+2Gf//sxb//68S//+sDv//qQv//6cI//+mCP//qAn//qsN/v/XP///1j7//9U9///VPP//1T3//9Y9///WPv//1z7//9c///7XP//5zzr/6Lw8/6TFoP94u7z/oKh8/5K4ov9bwNn/erSw/8qPI//Wjw3/6aca//m6Iv/+vSL//7of//+3G///sxb//7AS//+sDv//qQv//6gJ//+pCv//qw7//9hA///XP///1j7//9Y+///WPv//1z///9g////YQP/60Dv/7Lgr/9uZGP/Ap1X/Zdjr/1DM6v9Pw+L/Vtnz/1bd9v9Pv9v/i7Cc/6amcv+8nUv/2JMT/+qoGv/6uSH//rsg//+4HP//tBf//7AT//+tD///qw3//6sN//+tD///2UH//9hA///XP///1z///9hA//7YQP/3zjv/7bov/9+hHv/ZlBX/xaVQ/43G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
"@

	try {
		# Function to download and update all .ps1 scripts from a GitHub folder
		function Update-ScriptsFromGitHub {
			param(
				[Parameter(Mandatory)]
				[string]$GitHubRepo,  # Format: 'owner/repo'

				[Parameter(Mandatory)]
				[string]$GitHubFolder,  # Format: 'path/to/folder'

				[Parameter(Mandatory)]
				[string]$LocalFolder,

				[string]$Branch = 'master'
			)

			# Ensure local folder exists
			if (!(Test-Path $LocalFolder)) {
				New-Item -Path $LocalFolder -ItemType Directory -Force | Out-Null
			}

			try {
				# Suppress progress bar
				$oldProgressPreference = $ProgressPreference
				$ProgressPreference = 'SilentlyContinue'

				# Get folder contents from GitHub API using helper with caching
				$ps1Files = Get-GitHubFolderContents `
					-Owner ($GitHubRepo -split '/')[0] `
					-Repo ($GitHubRepo -split '/')[1] `
					-Path $GitHubFolder `
					-Branch $Branch `
					-FilePattern "*.ps1" `
					-UseCache

				foreach ($file in $ps1Files) {
					$localPath = Join-Path $LocalFolder $file.name

					# Download from raw URL
					$remoteContent = (Invoke-WebRequest -Uri $file.download_url -UseBasicParsing).Content

					# Normalize remote content to CRLF for Windows
					$remoteContentNormalized = $remoteContent -replace "`r`n", "`n"  # First normalize to LF
					$remoteContentNormalized = $remoteContentNormalized -replace "`n", "`r`n"  # Then convert to CRLF

					if (Test-Path $localPath) {
						$localContent = Get-Content $localPath -Raw -ErrorAction SilentlyContinue

						# Compare normalized content
						if ($remoteContentNormalized -eq $localContent) {
							continue
						}
					}

					# Save with CRLF line endings
					$remoteContentNormalized | Set-Content -Path $localPath -Encoding ASCII -NoNewline -Force
				}

				# Restore progress preference
				$ProgressPreference = $oldProgressPreference

			} catch {
				# Restore progress preference on error
				if ($oldProgressPreference) { $ProgressPreference = $oldProgressPreference }
				# Silent fail - fallback to local files
			}
		}

		# Function to check if current script is a default script
		function Test-IsDefaultScript {
			param([string]$FilePath)

			if ([string]::IsNullOrWhiteSpace($FilePath)) { return $false }

			$fileName = [System.IO.Path]::GetFileName($FilePath)
			return $fileName -in @('Std-Install.ps1', 'Std-Manifest.ps1', 'Std-WAU.ps1', 'Std-File.ps1')
		}

		# Function to load a default script (from disk if exists, otherwise hardcoded default)
		function Get-DefaultScriptContent {
			param(
				[string]$ScriptName,
				[string]$WsbDir
			)

			$scriptPath = Join-Path $WsbDir "$ScriptName.ps1"

			# Try to load from disk
			if (Test-Path $scriptPath) {
				try {
					return Get-Content -Path $scriptPath -Raw
				}
				catch {
					# Silently fail - return null
				}
			}

			# Return null if file doesn't exist
			return $null
		}


		# Initialize wsb directory and script mappings early
		# This also handles migration from old script names to new ones
		$wsbDir = Join-Path $Script:WorkingDir "wsb"
		Get-ScriptMappings | Out-Null  # Creates directory, mappings file, and migrates old names

		# Download/update default scripts from GitHub
		Write-Host "Checking default scripts...`t" -NoNewline -ForegroundColor Cyan
		$initialStatus = "Checking default scripts from GitHub"
		Update-ScriptsFromGitHub -GitHubRepo 'KnifMelti/SandboxStart' -GitHubFolder 'Source/assets/scripts' -LocalFolder $wsbDir
		Write-Host "Done" -ForegroundColor Green
		$initialStatus = "Default scripts ready"

		# Check for SandboxStart updates
		Write-Host "Checking for updates...`t`t" -NoNewline -ForegroundColor Cyan
		try {
			# Use GitHub API helper with caching
			$latestRelease = Get-GitHubLatestRelease `
				-Owner "KnifMelti" `
				-Repo "SandboxStart" `
				-UseCache

			# Get local SandboxStart.ps1 file timestamp
			$localScriptPath = Join-Path $Script:WorkingDir "SandboxStart.ps1"

			# Only check if the file exists
			if (Test-Path $localScriptPath) {
				# Get local file time in UTC for accurate comparison
				$localFileTime = (Get-Item $localScriptPath).LastWriteTimeUtc

				# Find the ZIP asset and use its created_at timestamp
				$zipAsset = $latestRelease.assets | Where-Object { $_.name -like "SandboxStart-*.zip" } | Select-Object -First 1
				if ($zipAsset) {
					# Parse ZIP asset created_at timestamp and convert to UTC (API returns UTC but Parse converts to local)
					$releaseDate = ([DateTime]::Parse($zipAsset.created_at)).ToUniversalTime()

					# Compare dates with 90 minute tolerance (build/upload time difference and timezone conversion)
					# If release is more than 90 minutes newer than local file, show update button
					$timeDifference = ($releaseDate - $localFileTime).TotalMinutes
					if ($timeDifference -gt 90) {
						$localFileOlderThanRelease = $true
						Write-Host "Update available" -ForegroundColor Yellow
					} else {
						Write-Host "Done" -ForegroundColor Green
					}
				} else {
					# No ZIP asset found
					Write-Host "Skipped" -ForegroundColor Gray
				}
			} else {
				# SandboxStart.ps1 doesn't exist (e.g., when running from shared submodule in other projects)
				Write-Host "Skipped" -ForegroundColor Gray
			}
		} catch {
			# Silent fail - don't show button if GitHub API is unreachable
			Write-Host "Skipped" -ForegroundColor Gray
		}

		# Load embedded icon
		try {
			$iconBytes = [System.Convert]::FromBase64String($iconBase64)
			$memoryStream = New-Object System.IO.MemoryStream($iconBytes, 0, $iconBytes.Length)
			$script:AppIcon = New-Object System.Drawing.Icon($memoryStream)
			$appIcon = $script:AppIcon
		}
		catch {
			Write-Warning "Failed to load embedded icon: $($_.Exception.Message)"
			$script:AppIcon = $null
			$appIcon = $null
		}

		# Create the main form
		$form = New-Object System.Windows.Forms.Form
		$form.Text = "Windows Sandbox Test Configuration"
		$form.Size = New-Object System.Drawing.Size(450, 725)
		$form.StartPosition = "CenterScreen"
		$form.FormBorderStyle = "FixedDialog"
		$form.MaximizeBox = $false
		$form.MinimizeBox = $false

		# Set custom icon if available
		if ($appIcon) {
			$form.Icon = $appIcon
			$form.ShowIcon = $true
		}
		else {
			$form.ShowIcon = $false
		}

		# Detect Windows theme preference (check if user has overridden it)
		if ($null -ne $script:UserThemeOverride) {
			$useDarkMode = $script:UserThemeOverride
		}
		else {
			$useDarkMode = Get-WindowsThemeSetting
		}

		# Define adaptive green color for Update button
		$updateButtonGreen = if ($useDarkMode) {
			[System.Drawing.Color]::FromArgb(60, 120, 60)  # Dark mode: muted green
		} else {
			[System.Drawing.Color]::LightGreen  # Light mode: bright green
		}

		# Create controls
		$y = 20
		$labelHeight = 20
		$controlHeight = 23
		$spacing = 5
		$leftMargin = 20
		$controlWidth = 400

		# Update button in top-right corner
		$btnUpdate = New-Object System.Windows.Forms.Button
		$btnUpdate.Name = 'btnUpdate'  # For theme detection
		$btnUpdate.Location = New-Object System.Drawing.Point($controlWidth, 15)
		$btnUpdate.Size = New-Object System.Drawing.Size(20, 20)
		$btnUpdate.Text = [char]0x2B06
		$btnUpdate.Font = New-Object System.Drawing.Font("Segoe UI Symbol", 14, [System.Drawing.FontStyle]::Regular)
		$btnUpdate.Visible = $false  # Initially hidden

		$toolTip = New-Object System.Windows.Forms.ToolTip
		$toolTip.SetToolTip($btnUpdate, "New version available - Click to download")

		$btnUpdate.Add_Click({
			Start-Process "https://github.com/KnifMelti/SandboxStart/releases/latest"
		})

		# Enable update button if newer version available
		if ($localFileOlderThanRelease) {
			$btnUpdate.Visible = $true
		}

		$form.Controls.Add($btnUpdate)

		# Mapped Folder selection
		$lblMapFolder = New-Object System.Windows.Forms.Label
		$lblMapFolder.Location = New-Object System.Drawing.Point($leftMargin, $y)
		$lblMapFolder.Size = New-Object System.Drawing.Size(150, $labelHeight)
		$lblMapFolder.Text = "Mapped Folder:"
		$form.Controls.Add($lblMapFolder)

		$txtMapFolder = New-Object System.Windows.Forms.TextBox
		$txtMapFolder.Location = New-Object System.Drawing.Point($leftMargin, ($y + $labelHeight))
		$txtMapFolder.Size = New-Object System.Drawing.Size($controlWidth, $controlHeight)
		# Set default path based on whether msi directory exists and find latest version
		$msiDir = Join-Path $Script:WorkingDir "msi"
		if (Test-Path $msiDir) {
			# Look for version directories (e.g., 2.6.1, 2.7.0) and get the latest one
			$versionDirs = Get-ChildItem -Path $msiDir -Directory | Where-Object { 
				$_.Name -match '^\d+\.\d+\.\d+$' 
			} | Sort-Object { [Version]$_.Name } -Descending
			
			if ($versionDirs) {
				$txtMapFolder.Text = $versionDirs[0].FullName
			} else {
				$txtMapFolder.Text = $msiDir
			}
		} else {
			$txtMapFolder.Text = $Script:WorkingDir
		}
		$form.Controls.Add($txtMapFolder)

		$y += $labelHeight + $controlHeight + 5

		# Folder browse button
		$btnBrowse = New-Object System.Windows.Forms.Button
		$btnBrowse.Location = New-Object System.Drawing.Point($leftMargin, $y)
		$btnBrowse.Size = New-Object System.Drawing.Size(($controlWidth / 2 - 5), $controlHeight)
		$btnBrowse.Text = "Folder..."
		$btnBrowse.Add_Click({
			# Use OpenFileDialog as folder picker for dark mode support
			$folderDialog = New-Object System.Windows.Forms.OpenFileDialog
			$folderDialog.ValidateNames = $false
			$folderDialog.CheckFileExists = $false
			$folderDialog.CheckPathExists = $true
			$folderDialog.FileName = "Select Folder"
			$folderDialog.Filter = "Folders|*.none"
			$folderDialog.Title = "Select folder to map in Windows Sandbox"

			# Set initial directory if path exists
			if (![string]::IsNullOrWhiteSpace($txtMapFolder.Text) -and (Test-Path $txtMapFolder.Text)) {
				$folderDialog.InitialDirectory = $txtMapFolder.Text
			}

			if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
				$selectedDir = Split-Path $folderDialog.FileName
				
				# Folder selected - use directory logic
				$txtMapFolder.Text = $selectedDir
				
				# Update sandbox folder name
				$msiFiles = Get-ChildItem -Path $selectedDir -Filter "WAU*.msi" -File -ErrorAction SilentlyContinue
				if ($msiFiles) {
					$txtSandboxFolderName.Text = "WAU-install"
				} else {
					$folderName = Split-Path $selectedDir -Leaf
					# Check if it's a root drive (contains : or is a path like D:\)
					if (![string]::IsNullOrWhiteSpace($folderName) -and $folderName -notmatch ':' -and $folderName -ne '\') {
						$txtSandboxFolderName.Text = $folderName
					} else {
						# Root drive selected (e.g., D:\) - use drive letter as folder name
						$driveLetter = $selectedDir.TrimEnd('\').Replace(':', '')
						if (![string]::IsNullOrWhiteSpace($driveLetter)) {
							$txtSandboxFolderName.Text = "Drive_$driveLetter"
						} else {
							$txtSandboxFolderName.Text = "MappedFolder"
						}
					}
				}
				
				# Find matching script from mappings
				$matchingScript = Find-MatchingScript -Path $selectedDir
				$scriptName = $matchingScript.Replace('.ps1', '')

				# Load script using the new dynamic loading function
				$scriptContent = Get-DefaultScriptContent -ScriptName $scriptName -WsbDir $wsbDir

				# Fallback to Installer if script not found
				if ([string]::IsNullOrWhiteSpace($scriptContent)) {
					$scriptContent = Get-DefaultScriptContent -ScriptName "Std-Install" -WsbDir $wsbDir
					$scriptName = "Std-Install"
					if ([string]::IsNullOrWhiteSpace($scriptContent)) {
						$lblStatus.Text = "Status: Scripts not available"
					} else {
						$lblStatus.Text = "Status: Mapping fallback to Std-Install.ps1"
					}
				} else {
					$lblStatus.Text = "Status: Mapping -> $matchingScript"
				}

				# Inject chosen folder name
				if ($scriptContent) {
					$scriptContent = $scriptContent -replace '\$SandboxFolderName\s*=\s*"[^"]*"', "`$SandboxFolderName = `"$($txtSandboxFolderName.Text)`""
					$txtScript.Text = $scriptContent

					# Update current script file tracking
					$script:currentScriptFile = Join-Path $wsbDir "$scriptName.ps1"

					# Update Save button state
					if (Test-IsDefaultScript -FilePath $script:currentScriptFile) {
						$btnSaveScript.Enabled = $false
					} else {
						$btnSaveScript.Enabled = $true
					}
				} else {
					$txtScript.Text = ""
					$script:currentScriptFile = $null
					$btnSaveScript.Enabled = $false
					$lblStatus.Text = "Status: Script not available"
				}
			}
		})
		$form.Controls.Add($btnBrowse)
		
		# File browse button
		$btnBrowseFile = New-Object System.Windows.Forms.Button
		$btnBrowseFile.Location = New-Object System.Drawing.Point(($leftMargin + $controlWidth / 2 + 5), $y)
		$btnBrowseFile.Size = New-Object System.Drawing.Size(($controlWidth / 2 - 5), $controlHeight)
		$btnBrowseFile.Text = "File..."
		$btnBrowseFile.Add_Click({
			$fileDialog = New-Object System.Windows.Forms.OpenFileDialog
			$fileDialog.Title = "Select file to run in Windows Sandbox"
			$fileDialog.Filter = "Executable Files (*.exe;*.msi;*.cmd;*.bat;*.ps1;*.ahk;*.py;*.js)|*.exe;*.msi;*.cmd;*.bat;*.ps1;*.ahk;*.py;*.js|All Files (*.*)|*.*"
			$fileDialog.InitialDirectory = $txtMapFolder.Text
			
			if ($fileDialog.ShowDialog() -eq "OK") {
				$selectedPath = $fileDialog.FileName
				$selectedDir = [System.IO.Path]::GetDirectoryName($selectedPath)
				$selectedFile = [System.IO.Path]::GetFileName($selectedPath)
				
				# File selected - use its directory
				$txtMapFolder.Text = $selectedDir
				
				# Update sandbox folder name based on directory only (no WAU detection)
				$folderName = Split-Path $selectedDir -Leaf
				# Check if it's a root drive (contains : or is a path like D:\)
				if (![string]::IsNullOrWhiteSpace($folderName) -and $folderName -notmatch ':' -and $folderName -ne '\') {
					$txtSandboxFolderName.Text = $folderName
				} else {
					# Root drive selected (e.g., D:\) - use drive letter as folder name
					$driveLetter = $selectedDir.TrimEnd('\').Replace(':', '')
					if (![string]::IsNullOrWhiteSpace($driveLetter)) {
						$txtSandboxFolderName.Text = "Drive_$driveLetter"
					} else {
						$txtSandboxFolderName.Text = "MappedFolder"
					}
				}
				
			# Generate script for selected file directly using Std-File.ps1
			$txtScript.Text = @"
`$SandboxFolderName = "$($txtSandboxFolderName.Text)"
& "`$env:USERPROFILE\Desktop\SandboxTest\Std-File.ps1" -SandboxFolderName `$SandboxFolderName -FileName "$selectedFile"
"@

			$lblStatus.Text = "Status: File selected -> $selectedFile (using Std-File.ps1)"

			# Auto-select Python package if .py file detected
			if ($selectedFile -like "*.py") {
				if ($chkNetworking.Checked) {
					# Check if Python package list exists
					$pythonPackageName = "Python"

					if ($cmbInstallPackages.Items -contains $pythonPackageName) {
						# Python package list exists - auto-select it
						$cmbInstallPackages.SelectedItem = $pythonPackageName
						$lblStatus.Text = "Status: .py selected -> Auto-selected Python package for installation"
					} else {
						# Python package list doesn't exist - show warning
						$lblStatus.Text = "Status: .py selected -> WARNING: create 'Python.txt' in wsb\ folder!"
					}
				} else {
					# Networking disabled - show warning
					$lblStatus.Text = "Status: .py selected -> WARNING: Enable networking (WinGet)!"
				}
			}

			# Auto-select AutoHotkey package if .ahk file detected
			if ($selectedFile -like "*.ahk") {
				if ($chkNetworking.Checked) {
					# Check if AutoHotkey package list exists
					$ahkPackageName = "AHK"

					if ($cmbInstallPackages.Items -contains $ahkPackageName) {
						# AHK package list exists - auto-select it
						$cmbInstallPackages.SelectedItem = $ahkPackageName
						$lblStatus.Text = "Status: .ahk selected -> Auto-selected AHK package for installation"
					} else {
						# AHK package list doesn't exist - show warning
						$lblStatus.Text = "Status: .ahk selected -> WARNING: create 'AHK.txt' in wsb\ folder!"
					}
				} else {
					# Networking disabled - show warning
					$lblStatus.Text = "Status: .ahk selected -> WARNING: Enable networking (WinGet)!"
				}
			}
			}
		})
		$form.Controls.Add($btnBrowseFile)

		$y += $labelHeight + $controlHeight + $spacing

		# Sandbox Folder Name
		$lblSandboxFolderName = New-Object System.Windows.Forms.Label
		$lblSandboxFolderName.Location = New-Object System.Drawing.Point($leftMargin, $y)
		$lblSandboxFolderName.Size = New-Object System.Drawing.Size(200, $labelHeight)
		$lblSandboxFolderName.Text = "Sandbox Desktop Folder Name:"
		$form.Controls.Add($lblSandboxFolderName)

		$txtSandboxFolderName = New-Object System.Windows.Forms.TextBox
		$txtSandboxFolderName.Location = New-Object System.Drawing.Point($leftMargin, ($y + $labelHeight))
		$txtSandboxFolderName.Size = New-Object System.Drawing.Size($controlWidth, $controlHeight)
		# Set default based on whether WAU MSI exists in the mapped folder
		$msiFiles = Get-ChildItem -Path $txtMapFolder.Text -Filter "WAU*.msi" -File -ErrorAction SilentlyContinue
		if ($msiFiles) {
			$txtSandboxFolderName.Text = "WAU-install"
		} else {
			$initialFolderName = Split-Path $txtMapFolder.Text -Leaf
			# Check if it's a root drive (contains : or is a path like D:\)
			if (![string]::IsNullOrWhiteSpace($initialFolderName) -and $initialFolderName -notmatch ':' -and $initialFolderName -ne '\') {
				$txtSandboxFolderName.Text = $initialFolderName
			} else {
				# Root drive - extract drive letter
				$driveLetter = $txtMapFolder.Text.TrimEnd('\').Replace(':', '')
				if (![string]::IsNullOrWhiteSpace($driveLetter)) {
					$txtSandboxFolderName.Text = "Drive_$driveLetter"
				} else {
					$txtSandboxFolderName.Text = "MappedFolder"
				}
			}
		}

		# Add event handler to update script when folder name changes
		$txtSandboxFolderName.Add_TextChanged({
			$currentScript = $txtScript.Text
			if (![string]::IsNullOrWhiteSpace($currentScript)) {
				# Replace the SandboxFolderName variable value in the existing script
				$txtScript.Text = $currentScript -replace '\$SandboxFolderName\s*=\s*"[^"]*"', "`$SandboxFolderName = `"$($txtSandboxFolderName.Text)`""
			}
		})

		$form.Controls.Add($txtSandboxFolderName)

		$y += $labelHeight + $controlHeight + $spacing

		# Install Packages section
		$lblInstallPackages = New-Object System.Windows.Forms.Label
		$lblInstallPackages.Location = New-Object System.Drawing.Point($leftMargin, $y)
		$lblInstallPackages.Size = New-Object System.Drawing.Size(200, $labelHeight)
		$lblInstallPackages.Text = "Install Packages:"
		$form.Controls.Add($lblInstallPackages)

		$cmbInstallPackages = New-Object System.Windows.Forms.ComboBox
		$cmbInstallPackages.Location = New-Object System.Drawing.Point($leftMargin, ($y + $labelHeight))
		$cmbInstallPackages.Size = New-Object System.Drawing.Size(($controlWidth - 85), $controlHeight)
		$cmbInstallPackages.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

		$tooltipPackages = New-Object System.Windows.Forms.ToolTip
		$tooltipPackages.SetToolTip($cmbInstallPackages, (Get-PackageListTooltip))

		# Populate dropdown
		$packageLists = Get-PackageLists
		[void]$cmbInstallPackages.Items.Add("")

		if ($packageLists.Count -eq 0) {
			$cmbInstallPackages.Items.Add("[Create new list...]")
			$cmbInstallPackages.SelectedIndex = 0
		} else {
			foreach ($list in $packageLists) {
				[void]$cmbInstallPackages.Items.Add($list)
			}
			[void]$cmbInstallPackages.Items.Add("[Create new list...]")
			$cmbInstallPackages.SelectedIndex = 0
		}

		# Selection change event
		$cmbInstallPackages.Add_SelectedIndexChanged({
			# Enable/disable Edit button based on selection
			$selectedItem = $this.SelectedItem
			$btnEditPackages.Enabled = ($selectedItem -ne "" -and $selectedItem -ne "[Create new list...]")

			if ($this.SelectedItem -eq "[Create new list...]") {
				$result = Show-PackageListEditor

				if ($result.DialogResult -eq 'OK') {
					$currentSelection = $result.ListName
					$this.Items.Clear()
					[void]$this.Items.Add("")

					$lists = Get-PackageLists
					foreach ($list in $lists) {
						[void]$this.Items.Add($list)
					}
					[void]$this.Items.Add("[Create new list...]")

					$this.SelectedItem = $currentSelection
					$tooltipPackages.SetToolTip($this, (Get-PackageListTooltip))
				} else {
					$this.SelectedIndex = 0
				}
			}
		})

		$form.Controls.Add($cmbInstallPackages)

		# Edit button
		$btnEditPackages = New-Object System.Windows.Forms.Button
		$btnEditPackages.Location = New-Object System.Drawing.Point(($leftMargin + $controlWidth - 75), ($y + $labelHeight))
		$btnEditPackages.Size = New-Object System.Drawing.Size(75, $controlHeight)
		$btnEditPackages.Text = "Edit..."
		$btnEditPackages.Enabled = $false  # Initially disabled until a valid package list is selected
		$btnEditPackages.Add_Click({
			$selectedList = $cmbInstallPackages.SelectedItem

			if ($selectedList -eq "" -or $selectedList -eq "[Create new list...]") {
				[System.Windows.Forms.MessageBox]::Show("Please select a package list to edit.", "No Selection", "OK", "Information")
				return
			}

			$result = Show-PackageListEditor -ListName $selectedList

			if ($result.DialogResult -eq 'OK') {
				$currentSelection = $selectedList
				$cmbInstallPackages.Items.Clear()
				[void]$cmbInstallPackages.Items.Add("")

				$lists = Get-PackageLists
				foreach ($list in $lists) {
					[void]$cmbInstallPackages.Items.Add($list)
				}
				[void]$cmbInstallPackages.Items.Add("[Create new list...]")

				$cmbInstallPackages.SelectedItem = $currentSelection
			}
		})
		$form.Controls.Add($btnEditPackages)

		$y += $labelHeight + $controlHeight + $spacing

		# WinGet Version - using ComboBox with fetched versions
		$lblWinGetVersion = New-Object System.Windows.Forms.Label
		$lblWinGetVersion.Location = New-Object System.Drawing.Point($leftMargin, $y)
		$lblWinGetVersion.Size = New-Object System.Drawing.Size(300, $labelHeight)
		$lblWinGetVersion.Text = "WinGet Version (leave empty for latest):"
		$form.Controls.Add($lblWinGetVersion)

		$cmbWinGetVersion = New-Object System.Windows.Forms.ComboBox
		$cmbWinGetVersion.Location = New-Object System.Drawing.Point($leftMargin, ($y + $labelHeight))
		$cmbWinGetVersion.Size = New-Object System.Drawing.Size($controlWidth, $controlHeight)
		$cmbWinGetVersion.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
	
	# Add empty option first (for "latest") - only item initially
	[void]$cmbWinGetVersion.Items.Add("")
		# Use Tag property to track if versions have been loaded (avoids script-scope issues)
		$cmbWinGetVersion.Tag = $false
		
		$cmbWinGetVersion.Add_DropDown({
			# Use $this to reference the ComboBox safely within the event handler
			if (-not $this.Tag) {
				# Show loading indicator by adding a temporary item
				$originalIndex = $this.SelectedIndex
				[void]$this.Items.Add("Loading versions...")
				$this.SelectedIndex = $this.Items.Count - 1
				[System.Windows.Forms.Application]::DoEvents()  # Force UI update
				
				try {
					Write-Verbose "Fetching stable WinGet versions for dropdown..."
					$stableVersions = Get-StableWinGetVersions
					
					# Remove loading indicator
					$this.Items.RemoveAt($this.Items.Count - 1)
					
					# Add fetched versions to the dropdown
					foreach ($version in $stableVersions) {
						[void]$this.Items.Add($version)
					}
					
					Write-Verbose "WinGet version dropdown populated with $($stableVersions.Count) stable versions"
				}
				catch {
					Write-Warning "Failed to populate WinGet versions dropdown: $($_.Exception.Message)"
					# Remove loading indicator even on error
					if ($this.Items.Count -gt 0 -and $this.Items[$this.Items.Count - 1] -eq "Loading versions...") {
						$this.Items.RemoveAt($this.Items.Count - 1)
					}
				}
				finally {
					# Restore original selection and mark as loaded
					$this.SelectedIndex = $originalIndex
					$this.Tag = $true
				}
			}
		})
		
		$form.Controls.Add($cmbWinGetVersion)

		$y += $labelHeight + $controlHeight + $spacing + 10

		# Checkboxes
		$chkPrerelease = New-Object System.Windows.Forms.CheckBox
		$chkPrerelease.Location = New-Object System.Drawing.Point($leftMargin, $y)
		$chkPrerelease.Size = New-Object System.Drawing.Size(200, $labelHeight)
		$chkPrerelease.Text = "Pre-release (of WinGet)"
		$form.Controls.Add($chkPrerelease)
		
		# Add event handler after both controls are added to form
		# Store reference to combo box in checkbox's Tag for safe access
		$chkPrerelease.Tag = $cmbWinGetVersion
		$chkPrerelease.Add_CheckedChanged({
			$comboBox = $this.Tag
			if ($this.Checked) {
				# Disable version field when Pre-release is checked
				$comboBox.Enabled = $false
				$comboBox.Text = ""
			} else {
				# Enable version field when Pre-release is unchecked (but only if networking is enabled)
				$comboBox.Enabled = $chkNetworking.Checked
			}
		})

		$y += $labelHeight + 5

		$chkClean = New-Object System.Windows.Forms.CheckBox
		$chkClean.Location = New-Object System.Drawing.Point($leftMargin, $y)
		$chkClean.Size = New-Object System.Drawing.Size(200, $labelHeight)
		$chkClean.Text = "Clean (cached dependencies)"
		$form.Controls.Add($chkClean)

		$y += $labelHeight + 5

		$chkVerbose = New-Object System.Windows.Forms.CheckBox
		$chkVerbose.Location = New-Object System.Drawing.Point($leftMargin, $y)
		$chkVerbose.Size = New-Object System.Drawing.Size(300, $labelHeight)
		$chkVerbose.Text = "Verbose (screen log && wait)"
		$form.Controls.Add($chkVerbose)

		$y += $labelHeight + $spacing + 10

		# WSB Configuration Section
		$lblWSBConfig = New-Object System.Windows.Forms.Label
		$lblWSBConfig.Location = New-Object System.Drawing.Point($leftMargin, $y)
		$lblWSBConfig.Size = New-Object System.Drawing.Size(300, $labelHeight)
		$lblWSBConfig.Text = "WSB Configuration:"
		$lblWSBConfig.Font = New-Object System.Drawing.Font($lblWSBConfig.Font.FontFamily, $lblWSBConfig.Font.Size, [System.Drawing.FontStyle]::Bold)
		$form.Controls.Add($lblWSBConfig)

		$y += $labelHeight + 5

		# Networking checkbox
		$chkNetworking = New-Object System.Windows.Forms.CheckBox
		$chkNetworking.Location = New-Object System.Drawing.Point($leftMargin, $y)
		$chkNetworking.Size = New-Object System.Drawing.Size(200, $labelHeight)
		$chkNetworking.Text = "Enable Networking"
		$chkNetworking.Checked = $true
		$tooltipNetworking = New-Object System.Windows.Forms.ToolTip
		$tooltipNetworking.SetToolTip($chkNetworking, "Enable network access in sandbox (required for WinGet downloads)")

		# Add event handler to enable/disable WinGet-related controls based on networking
		$chkNetworking.Add_CheckedChanged({
			$enabled = $this.Checked

			# Enable/disable all WinGet-related controls
			$cmbInstallPackages.Enabled = $enabled
			# Edit button requires both networking enabled AND a valid package list selected
			$selectedItem = $cmbInstallPackages.SelectedItem
			$btnEditPackages.Enabled = $enabled -and ($selectedItem -ne "" -and $selectedItem -ne "[Create new list...]")
			$cmbWinGetVersion.Enabled = $enabled -and -not $chkPrerelease.Checked
			$chkPrerelease.Enabled = $enabled
			$chkClean.Enabled = $enabled

			# Clear selections when disabling
			if (-not $enabled) {
				$cmbInstallPackages.SelectedIndex = 0  # Select empty option
				$cmbWinGetVersion.SelectedIndex = 0    # Select empty option
				$chkPrerelease.Checked = $false
				$chkClean.Checked = $false
			}
		})

		$form.Controls.Add($chkNetworking)

		$y += $labelHeight + 5

		# Memory dropdown - First detect available system RAM
		# Get total physical memory and calculate safe maximum (75% of total RAM)
		# Uses multiple methods (no elevation required):
		# 1. ComputerInfo (Win10+, preferred - fastest and no CIM)
		# 2. Win32_ComputerSystem via Get-CimInstance (fallback)
		# 3. WMI via Get-WmiObject (older systems)
		# 4. Hard-coded fallback (8 GB)
		Write-Host "Detecting system memory...`t" -NoNewline -ForegroundColor Cyan
		try {
			$totalMemoryMB = $null

			# Method 1: Try ComputerInfo (Windows 10+ preferred method - no CIM/WMI)
			try {
				# Suppress progress output from Get-ComputerInfo
				$prevProgressPreference = $ProgressPreference
				$ProgressPreference = 'SilentlyContinue'
				$computerInfo = Get-ComputerInfo -Property CsTotalPhysicalMemory -ErrorAction Stop
				$ProgressPreference = $prevProgressPreference
				$totalMemoryMB = [int]($computerInfo.CsTotalPhysicalMemory / 1MB)
				Write-Host "Done" -ForegroundColor Green
				Write-Verbose "Memory detected via ComputerInfo: $totalMemoryMB MB"
			}
			catch {
				# Restore progress preference on error
				if ($prevProgressPreference) { $ProgressPreference = $prevProgressPreference }
				# Method 2: Try CIM (works without elevation on most systems)
				try {
					$totalMemoryMB = [int]((Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop).TotalPhysicalMemory / 1MB)
					Write-Host "Done" -ForegroundColor Green
					Write-Verbose "Memory detected via CIM: $totalMemoryMB MB"
				}
				catch {
					# Method 3: Try WMI as last resort
					try {
						$totalMemoryMB = [int]((Get-WmiObject -Class Win32_ComputerSystem -ErrorAction Stop).TotalPhysicalMemory / 1MB)
						Write-Host "Done" -ForegroundColor Green
						Write-Verbose "Memory detected via WMI: $totalMemoryMB MB"
					}
					catch {
						# All methods failed - use fallback
						$totalMemoryMB = $null
					}
				}
			}

			if ($totalMemoryMB) {
				$maxSafeMemory = [int]($totalMemoryMB * 0.75)
			}
			else {
				throw "All memory detection methods failed"
			}
		}
		catch {
			# Fallback if all detection methods fail
			Write-Host " Using default" -ForegroundColor Yellow
			$totalMemoryMB = 8192
			$maxSafeMemory = 6144
			Write-Verbose "Could not detect system memory, using fallback: $totalMemoryMB MB (max safe: $maxSafeMemory MB)"
		}

		# Generate memory options dynamically based on available RAM
		# Start with common increments and filter based on maxSafeMemory
		$allMemoryOptions = @(2048, 4096, 6144, 8192, 10240, 12288, 16384, 20480, 24576, 32768, 49152, 65536)
		$memoryOptions = $allMemoryOptions | Where-Object { $_ -le $maxSafeMemory }

		# Ensure minimum option exists (2048 MB required by Windows Sandbox)
		if (-not $memoryOptions -or $memoryOptions.Count -eq 0) {
			$memoryOptions = @(2048)
		}

		$defaultSelected = -1

		# Now create the UI controls with the detected memory values
		$lblMemory = New-Object System.Windows.Forms.Label
		$lblMemory.Location = New-Object System.Drawing.Point($leftMargin, $y)
		$lblMemory.Size = New-Object System.Drawing.Size(120, $labelHeight)
		$lblMemory.Text = "Memory (MB):"
		$form.Controls.Add($lblMemory)

		$cmbMemory = New-Object System.Windows.Forms.ComboBox
		$cmbMemory.Location = New-Object System.Drawing.Point(($leftMargin + 139), $y)
		$cmbMemory.Size = New-Object System.Drawing.Size(120, $controlHeight)
		$cmbMemory.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

		# Populate with pre-calculated memory options
		foreach ($memOption in $memoryOptions) {
			if ($memOption -le $maxSafeMemory) {
				[void]$cmbMemory.Items.Add($memOption.ToString())
				if ($memOption -eq 4096) { $defaultSelected = $cmbMemory.Items.Count - 1 }
			}
		}

		# Set default selection (4096 MB if available, otherwise the highest safe option)
		if ($defaultSelected -ge 0) {
			$cmbMemory.SelectedIndex = $defaultSelected
		}
		elseif ($cmbMemory.Items.Count -gt 0) {
			$cmbMemory.SelectedIndex = $cmbMemory.Items.Count - 1  # Select highest available
		}
		else {
			# Extreme fallback: add minimum required memory
			[void]$cmbMemory.Items.Add("2048")
			$cmbMemory.SelectedIndex = 0
		}

		# Build helpful tooltip showing available RAM and highest option
		$highestOption = if ($memoryOptions.Count -gt 0) { $memoryOptions[-1] } else { 2048 }
		$highestOptionGB = [math]::Round($highestOption / 1024, 1)
		$totalGB = [math]::Round($totalMemoryMB / 1024, 1)

		$tooltipMemory = New-Object System.Windows.Forms.ToolTip
		$tooltipMemory.SetToolTip($cmbMemory, "RAM for sandbox. Your system: $totalGB GB total. Highest safe option: $highestOptionGB GB (leaves 25% for Windows)")
		$form.Controls.Add($cmbMemory)

		$y += $labelHeight + 5

		# vGPU dropdown
		$lblvGPU = New-Object System.Windows.Forms.Label
		$lblvGPU.Location = New-Object System.Drawing.Point($leftMargin, $y)
		$lblvGPU.Size = New-Object System.Drawing.Size(120, $labelHeight)
		$lblvGPU.Text = "GPU Virtualization:"
		$form.Controls.Add($lblvGPU)

		$cmbvGPU = New-Object System.Windows.Forms.ComboBox
		$cmbvGPU.Location = New-Object System.Drawing.Point(($leftMargin + 139), $y)
		$cmbvGPU.Size = New-Object System.Drawing.Size(120, $controlHeight)
		$cmbvGPU.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
		[void]$cmbvGPU.Items.Add("Default")
		[void]$cmbvGPU.Items.Add("Enable")
		[void]$cmbvGPU.Items.Add("Disable")
		$cmbvGPU.SelectedIndex = 0  # Default: "Default"
		$tooltipvGPU = New-Object System.Windows.Forms.ToolTip
		$tooltipvGPU.SetToolTip($cmbvGPU, "Enable: Hardware acceleration, Disable: Software rendering (WARP), Default: System default (currently Enable)")
		$form.Controls.Add($cmbvGPU)

		$y += $labelHeight + $spacing + 10

		# Track the currently edited script file path for Save button
		$script:currentScriptFile = $null

		# Script section
		# Clear script button (X)
		$btnClearScript = New-Object System.Windows.Forms.Button
		$btnClearScript.Location = New-Object System.Drawing.Point(($leftMargin + $controlWidth - 20), ($y + 3))
		$btnClearScript.Size = New-Object System.Drawing.Size(20, 20)
		$btnClearScript.Text = [char]0x2716
		$btnClearScript.Font = New-Object System.Drawing.Font("Segoe UI Symbol", 14, [System.Drawing.FontStyle]::Regular)

		# Tooltip
		$tooltipClearScript = New-Object System.Windows.Forms.ToolTip
		$tooltipClearScript.SetToolTip($btnClearScript, "Clear editor")

		# Click event
		$btnClearScript.Add_Click({
			$txtScript.Text = ""
			$script:currentScriptFile = $null
			$btnSaveScript.Enabled = $false
			$lblStatus.Text = "Status: Script editor cleared"
		})

		$form.Controls.Add($btnClearScript)

		# Edit mappings button (pen icon)
		$btnEditMappings = New-Object System.Windows.Forms.Button
		$btnEditMappings.Location = New-Object System.Drawing.Point(20, ($y + 3))
		$btnEditMappings.Size = New-Object System.Drawing.Size(20, 20)
		$btnEditMappings.Text = [char]0x270E
		$btnEditMappings.Font = New-Object System.Drawing.Font("Segoe UI Symbol", 14, [System.Drawing.FontStyle]::Regular)

		# Tooltip
		$tooltipEditMappings = New-Object System.Windows.Forms.ToolTip
		$tooltipEditMappings.SetToolTip($btnEditMappings, "Edit mappings...")

		# Click event
		$btnEditMappings.Add_Click({
			Show-PackageListEditor -EditorMode "ScriptMapping" | Out-Null
		})

		$form.Controls.Add($btnEditMappings)

		# Create the script textbox first (before buttons) so buttons appear on top
		$txtScript = New-Object System.Windows.Forms.TextBox
		$txtScript.Location = New-Object System.Drawing.Point($leftMargin, ($y + $labelHeight + 5))
		$txtScript.Size = New-Object System.Drawing.Size($controlWidth, 120)
		$txtScript.Multiline = $true
		$txtScript.ScrollBars = "Vertical"
		$txtScript.AcceptsReturn = $true

		# Tooltip for script editor
		$tooltipScript = New-Object System.Windows.Forms.ToolTip
		$tooltipScript.SetToolTip($txtScript, "User script that will be executed in Windows Sandbox")

		# Set default script based on folder contents
		try {
			$installWSBPath = Join-Path $txtMapFolder.Text "InstallWSB.cmd"
			$installerYamlFiles = Get-ChildItem -Path $txtMapFolder.Text -Filter "*.installer.yaml" -File -ErrorAction SilentlyContinue
			# Use mapping on initial folder to detect Installer.ps1 scenario
			$matchingScriptInit = Find-MatchingScript -Path $txtMapFolder.Text

			# Determine which script to load based on folder contents
			$selectedScriptName = $null
			$autoDetectedStatus = ""
			if (Test-Path $installWSBPath) {
				$selectedScriptName = "Std-WAU"
				$script:currentScriptFile = Join-Path $wsbDir "Std-WAU.ps1"
				$autoDetectedStatus = "Auto-loaded: Std-WAU.ps1 (InstallWSB.cmd found)"
			} elseif ($installerYamlFiles) {
				$selectedScriptName = "Std-Manifest"
				$script:currentScriptFile = Join-Path $wsbDir "Std-Manifest.ps1"
				$autoDetectedStatus = "Auto-loaded: Std-Manifest.ps1 (*.installer.yaml found)"
			} elseif ($matchingScriptInit) {
				# Use whatever script the mapping system returned
				$selectedScriptName = $matchingScriptInit.Replace('.ps1', '')
				$script:currentScriptFile = Join-Path $wsbDir $matchingScriptInit
				if ($matchingScriptInit -eq 'Std-Install.ps1') {
					$autoDetectedStatus = "Auto-loaded: Std-Install.ps1 (default)"
				} else {
					$autoDetectedStatus = "Auto-loaded: $matchingScriptInit (from mapping)"
				}
			} else {
				# True fallback - only if Find-MatchingScript returns nothing
				$selectedScriptName = "Std-Install"
				$script:currentScriptFile = Join-Path $wsbDir "Std-Install.ps1"
				$autoDetectedStatus = "Auto-loaded: Std-Install.ps1 (default)"
			}

			# Only override initialStatus if it's still the "scripts ready" message
			if ($initialStatus -eq "Default scripts ready") {
				$initialStatus = $autoDetectedStatus
			}

			# Load the selected script using dynamic loading
			$scriptContent = Get-DefaultScriptContent -ScriptName $selectedScriptName -WsbDir $wsbDir
			if ($scriptContent) {
				$txtScript.Text = ($scriptContent -replace '\$SandboxFolderName\s*=\s*"[^"]*"', "`$SandboxFolderName = `"$($txtSandboxFolderName.Text)`"")
			} else {
				$txtScript.Text = ""
				# Override status to show script is missing
				$initialStatus = "Warning: $selectedScriptName.ps1 not available"
			}
		}
		catch {
			# Fallback to Installer script if anything goes wrong
			$selectedScriptName = "Std-Install"
			$script:currentScriptFile = Join-Path $wsbDir "Std-Install.ps1"
			if ($initialStatus -eq "Default scripts ready") {
				$initialStatus = "Auto-loaded: Std-Install.ps1 (default - error during detection)"
			}
			$scriptContent = Get-DefaultScriptContent -ScriptName "Std-Install" -WsbDir $wsbDir
			if ($scriptContent) {
				$txtScript.Text = $scriptContent
			} else {
				$txtScript.Text = ""
				# Override status to show script is missing
				$initialStatus = "Warning: Std-Install.ps1 not available"
			}
		}
		$form.Controls.Add($txtScript)

		# Load/Save buttons for scripts (added after txtScript so they appear on top)
		$btnLoadScript = New-Object System.Windows.Forms.Button
		$btnLoadScript.Location = New-Object System.Drawing.Point(103, $y)
		$btnLoadScript.Size = New-Object System.Drawing.Size(75, $controlHeight)
		$btnLoadScript.Text = "Load..."
		$btnLoadScript.Add_Click({
			$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
			$openFileDialog.InitialDirectory = $wsbDir
			$openFileDialog.Filter = "PowerShell Scripts (*.ps1)|*.ps1"
			$openFileDialog.Title = "Load Script"

			if ($openFileDialog.ShowDialog() -eq "OK") {
				try {
					$scriptContent = Get-Content -Path $openFileDialog.FileName -Raw
					$txtScript.Text = $scriptContent

					# Track the loaded file for the Save button
					$script:currentScriptFile = $openFileDialog.FileName

					# Extract SandboxFolderName from the loaded script
					$pattern = '\$SandboxFolderName\s*=\s*"([^"]*)"'
					if ($scriptContent -match $pattern) {
						$extractedFolderName = $matches[1]
						if (![string]::IsNullOrWhiteSpace($extractedFolderName) -and $extractedFolderName -ne "DefaultFolder") {
							$txtSandboxFolderName.Text = $extractedFolderName
						}
					}

					# Update script content with current folder name from the text field
					$currentFolderName = $txtSandboxFolderName.Text
					if (![string]::IsNullOrWhiteSpace($currentFolderName)) {
						$txtScript.Text = $txtScript.Text -replace '\$SandboxFolderName\s*=\s*"[^"]*"', "`$SandboxFolderName = `"$currentFolderName`""
					}

					# Update Save button state for loaded file
					if (Test-IsDefaultScript -FilePath $script:currentScriptFile) {
						$btnSaveScript.Enabled = $false
					} else {
						$btnSaveScript.Enabled = $true
					}

					# Update status to show loaded script
					$scriptFileName = [System.IO.Path]::GetFileName($openFileDialog.FileName)
					$lblStatus.Text = "Status: Loaded $scriptFileName"
				}
				catch {
					[System.Windows.Forms.MessageBox]::Show("Error loading script: $($_.Exception.Message)", "Load Error", "OK", "Error")
					$lblStatus.Text = "Status: Error loading script"
				}
			}
		})
		$form.Controls.Add($btnLoadScript)

		$btnSaveScript = New-Object System.Windows.Forms.Button
		$btnSaveScript.Location = New-Object System.Drawing.Point(183, $y)
		$btnSaveScript.Size = New-Object System.Drawing.Size(75, $controlHeight)
		$btnSaveScript.Text = "Save"
		$btnSaveScript.Add_Click({
			if ([string]::IsNullOrWhiteSpace($txtScript.Text)) {
				[System.Windows.Forms.MessageBox]::Show("No script content to save.", "Save Error", "OK", "Warning")
				return
			}

			# If we have a current file, save directly
			if ($script:currentScriptFile -and (Test-Path (Split-Path $script:currentScriptFile -Parent))) {
				try {
					$txtScript.Text | Out-File -FilePath $script:currentScriptFile -Encoding ASCII
					[System.Windows.Forms.MessageBox]::Show("Script saved successfully to:`n$($script:currentScriptFile)", "Save Complete", "OK", "Information")
				}
				catch {
					[System.Windows.Forms.MessageBox]::Show("Error saving script: $($_.Exception.Message)", "Save Error", "OK", "Error")
				}
			}
			else {
				# No current file or parent directory doesn't exist - show Save As dialog
				$saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
				$saveFileDialog.InitialDirectory = $wsbDir
				$saveFileDialog.Filter = "PowerShell Scripts (*.ps1)|*.ps1"
				$saveFileDialog.DefaultExt = "ps1"
				$saveFileDialog.Title = "Save Script"

				if ($saveFileDialog.ShowDialog() -eq "OK") {
					# Enforce .ps1 extension
					$targetPath = if ([System.IO.Path]::GetExtension($saveFileDialog.FileName).ToLower() -ne ".ps1") {
						"$($saveFileDialog.FileName).ps1"
					} else {
						$saveFileDialog.FileName
					}

					try {
						# Ensure wsb directory exists
						if (-not (Test-Path $wsbDir)) {
							New-Item -ItemType Directory -Path $wsbDir -Force | Out-Null
						}
						$txtScript.Text | Out-File -FilePath $targetPath -Encoding ASCII

						# Update current file tracking
						$script:currentScriptFile = $targetPath

						[System.Windows.Forms.MessageBox]::Show("Script saved successfully to:`n$targetPath", "Save Complete", "OK", "Information")
					}
					catch {
						[System.Windows.Forms.MessageBox]::Show("Error saving script: $($_.Exception.Message)", "Save Error", "OK", "Error")
					}
				}
			}
		})
		$form.Controls.Add($btnSaveScript)

		# Set initial Save button state based on current script
		if ($script:currentScriptFile -and (Test-IsDefaultScript -FilePath $script:currentScriptFile)) {
			$btnSaveScript.Enabled = $false
		} elseif ([string]::IsNullOrWhiteSpace($txtScript.Text)) {
			$btnSaveScript.Enabled = $false
		}

		# Add TextChanged event to update Save button state dynamically
		$txtScript.Add_TextChanged({
			if (Test-IsDefaultScript -FilePath $script:currentScriptFile) {
				# Default scripts cannot be saved
				$btnSaveScript.Enabled = $false
			} elseif ([string]::IsNullOrWhiteSpace($txtScript.Text)) {
				# Empty script cannot be saved
				$btnSaveScript.Enabled = $false
			} elseif ([string]::IsNullOrWhiteSpace($script:currentScriptFile)) {
				# No file path - must use Save As
				$btnSaveScript.Enabled = $false
			} else {
				# Valid non-default script with content
				$btnSaveScript.Enabled = $true
			}
		})


		$btnSaveAsScript = New-Object System.Windows.Forms.Button
		$btnSaveAsScript.Location = New-Object System.Drawing.Point(263, $y)
		$btnSaveAsScript.Size = New-Object System.Drawing.Size(75, $controlHeight)
		$btnSaveAsScript.Text = "Save as..."
		$btnSaveAsScript.Add_Click({
			if ([string]::IsNullOrWhiteSpace($txtScript.Text)) {
				[System.Windows.Forms.MessageBox]::Show("No script content to save.", "Save Error", "OK", "Warning")
				return
			}

			$saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
			$saveFileDialog.InitialDirectory = $wsbDir
			$saveFileDialog.Filter = "PowerShell Scripts (*.ps1)|*.ps1"
			$saveFileDialog.DefaultExt = "ps1"
			$saveFileDialog.Title = "Save Script As"

			# Pre-populate filename only if it's NOT a default script
			if ($script:currentScriptFile -and -not (Test-IsDefaultScript -FilePath $script:currentScriptFile)) {
				$saveFileDialog.FileName = [System.IO.Path]::GetFileName($script:currentScriptFile)
			}

			if ($saveFileDialog.ShowDialog() -eq "OK") {
				# Enforce .ps1 extension
				$targetPath = if ([System.IO.Path]::GetExtension($saveFileDialog.FileName).ToLower() -ne ".ps1") {
					"$($saveFileDialog.FileName).ps1"
				} else {
					$saveFileDialog.FileName
				}

				try {
					# Ensure wsb directory exists
					if (-not (Test-Path $wsbDir)) {
						New-Item -ItemType Directory -Path $wsbDir -Force | Out-Null
					}
					$txtScript.Text | Out-File -FilePath $targetPath -Encoding ASCII

					# Update current file tracking
					$script:currentScriptFile = $targetPath

					[System.Windows.Forms.MessageBox]::Show("Script saved successfully to:`n$targetPath", "Save Complete", "OK", "Information")
				}
				catch {
					[System.Windows.Forms.MessageBox]::Show("Error saving script: $($_.Exception.Message)", "Save Error", "OK", "Error")
				}
			}
		})
		$form.Controls.Add($btnSaveAsScript)

		# Status label (mapping/result info)
		$y += $labelHeight + 5 + 120 + 5
		$lblStatus = New-Object System.Windows.Forms.Label
		$lblStatus.Location = New-Object System.Drawing.Point($leftMargin, $y)
		$lblStatus.Size = New-Object System.Drawing.Size($controlWidth, $labelHeight)
		$lblStatus.Text = "Status: $initialStatus"
		$form.Controls.Add($lblStatus)

		$y += $labelHeight + $spacing + 10

		# Buttons
		$btnOK = New-Object System.Windows.Forms.Button
		$btnOK.Location = New-Object System.Drawing.Point(($leftMargin + $controlWidth - 160), $y)
		$btnOK.Size = New-Object System.Drawing.Size(75, 30)
		$btnOK.Text = "OK"
		$btnOK.Add_Click({
			$resultScript = $null
			if (-not [string]::IsNullOrWhiteSpace($txtScript.Text)) {
				try { $resultScript = [ScriptBlock]::Create($txtScript.Text) } catch { $resultScript = $null }
			}

			$script:__dialogReturn = @{
				DialogResult = 'OK'
				MapFolder = $txtMapFolder.Text
				SandboxFolderName = $txtSandboxFolderName.Text
				WinGetVersion = $cmbWinGetVersion.Text
				InstallPackageList = if ($cmbInstallPackages.SelectedItem -and
				                       $cmbInstallPackages.SelectedItem -ne "" -and
				                       $cmbInstallPackages.SelectedItem -ne "[Create new list...]") {
				                       $cmbInstallPackages.SelectedItem
				                   } else { "" }
				Prerelease = $chkPrerelease.Checked
				Clean = $chkClean.Checked
				Verbose = $chkVerbose.Checked
				Wait = $chkVerbose.Checked
				Networking = if ($chkNetworking.Checked) { "Enable" } else { "Disable" }
				MemoryInMB = [int]$cmbMemory.SelectedItem
				vGPU = $cmbvGPU.SelectedItem
				Script = $resultScript
			}
			$form.Close()
		})
		$form.Controls.Add($btnOK)

		$btnCancel = New-Object System.Windows.Forms.Button
		$btnCancel.Location = New-Object System.Drawing.Point(($leftMargin + $controlWidth - 75), $y)
		$btnCancel.Size = New-Object System.Drawing.Size(75, 30)
		$btnCancel.Text = "Cancel"
		$btnCancel.Add_Click({
			$script:__dialogReturn = @{ DialogResult = 'Cancel' }
			$form.Close()
		})
		$form.Controls.Add($btnCancel)

		# Set default accept/cancel buttons
		$form.AcceptButton = $btnOK
		$form.CancelButton = $btnCancel

		# Add double-click event to toggle dark/light mode
		$form.Add_DoubleClick({
			& $global:ToggleFormThemeHandler $form $updateButtonGreen
		}.GetNewClosure())

		# Apply theme based on Windows settings
		if ($useDarkMode) {
			Set-DarkModeTheme -Control $form -UpdateButtonBackColor $updateButtonGreen
			Set-DarkTitleBar -Form $form -UseDarkMode $true
		}
		else {
			Set-LightModeTheme -Control $form -UpdateButtonBackColor $updateButtonGreen
			Set-DarkTitleBar -Form $form -UseDarkMode $false
		}

		# Show dialog (modal)
		[void]$form.ShowDialog()

		# Prepare return object
		if ($script:__dialogReturn) {
			return $script:__dialogReturn
		} else {
			return @{ DialogResult = 'Cancel' }
		}
	}
	catch {
		$errorMsg = "Error creating dialog: $($_.Exception.Message)`n`nStack Trace:`n$($_.ScriptStackTrace)"
		Write-Host $errorMsg
		[System.Windows.Forms.MessageBox]::Show($errorMsg, "Error", "OK", "Error")
		return @{ DialogResult = "Cancel" }
	}
	finally {
		if ($form) { $form.Dispose() }
		if ($appIcon) { $appIcon.Dispose() }
		if ($memoryStream) { $memoryStream.Dispose() }
	}
}

# Show configuration dialog
$dialogResult = Show-SandboxTestDialog

if ($dialogResult.DialogResult -ne 'OK') {
	# User cancelled the dialog
	exit
}

# Build parameters for SandboxTest
$sandboxParams = @{
	MapFolder = $dialogResult.MapFolder
	SandboxFolderName = $dialogResult.SandboxFolderName
	Script = $dialogResult.Script
}

# Add optional parameters if they have values
if (![string]::IsNullOrWhiteSpace($dialogResult.WinGetVersion)) {
	$sandboxParams.WinGetVersion = $dialogResult.WinGetVersion
}
if (![string]::IsNullOrWhiteSpace($dialogResult.InstallPackageList)) {
	$sandboxParams.InstallPackageList = $dialogResult.InstallPackageList
}
if ($dialogResult.Prerelease) { $sandboxParams.Prerelease = $true }
if ($dialogResult.Clean) { $sandboxParams.Clean = $true }
$sandboxParams.Async = $true
if ($dialogResult.Verbose) { $sandboxParams.Verbose = $true }

# Add WSB configuration parameters
if (![string]::IsNullOrWhiteSpace($dialogResult.Networking)) {
	$sandboxParams.Networking = $dialogResult.Networking
}
if ($dialogResult.MemoryInMB) {
	$sandboxParams.MemoryInMB = $dialogResult.MemoryInMB
}
if (![string]::IsNullOrWhiteSpace($dialogResult.vGPU)) {
	$sandboxParams.vGPU = $dialogResult.vGPU
}

# Call SandboxTest with collected parameters
SandboxTest @sandboxParams

# Wait for key press if requested
if ($dialogResult.Wait) {
	Write-Host "`nPress any key to exit..." -ForegroundColor Yellow
	$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

exit

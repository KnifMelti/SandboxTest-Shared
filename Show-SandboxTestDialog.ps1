# Note: Theme preference is now stored in registry via Get-SandboxStartThemePreference()
# Old $script:UserThemeOverride variable removed - use registry functions instead

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
		Set-Content -Path $mappingFile -Value $defaultContent -Encoding UTF8
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
			Set-Content -Path $mappingFile -Value ($updatedLines -join "`r`n") -Encoding UTF8

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
	Filters out deleted lists (state=0 in package-lists.ini) and ensures
	AutoInstall appears first.

	.OUTPUTS
	Array of package list names (without .txt extension)
	#>

	$packageListDir = Join-Path $Script:WorkingDir "wsb"
	$config = Get-SandboxConfig -Section 'Lists' -WorkingDir $Script:WorkingDir
	$lists = @()
	$hasAutoInstall = $false

	if (Test-Path $packageListDir) {
		$txtFiles = Get-ChildItem -Path $packageListDir -Filter "*.txt" -File -ErrorAction SilentlyContinue
		foreach ($file in $txtFiles) {
			# Exclude script-mappings.txt from package lists
			if ($file.Name -ne "script-mappings.txt") {
				$listName = $file.BaseName

				# Filter out deleted lists (state=0 in .ini)
				if ($config.ContainsKey($listName) -and $config[$listName] -eq 0) {
					continue
				}

				if ($listName -eq "AutoInstall") {
					$hasAutoInstall = $true
				} else {
					$lists += $listName
				}
			}
		}
	}

	# Sort regular lists, but prepend AutoInstall
	$sortedLists = $lists | Sort-Object
	if ($hasAutoInstall) {
		return @('AutoInstall') + $sortedLists
	} else {
		return $sortedLists
	}
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
		[ValidateSet("PackageList", "ScriptMapping", "ConfigEdit")]
		[string]$EditorMode = "PackageList"
	)

	# Create editor form
	$editorForm = New-Object System.Windows.Forms.Form
	$editorForm.Text = switch ($EditorMode) {
		"PackageList" { if ($ListName) { "Edit Package List: $ListName" } else { "Create New Package List" } }
		"ScriptMapping" { "Edit Script Mappings" }
		"ConfigEdit" { "Edit Configuration" }
	}
	$editorForm.Size = switch ($EditorMode) {
		"PackageList" { New-Object System.Drawing.Size(420, 370) }
		"ScriptMapping" { New-Object System.Drawing.Size(510, 505) }
		"ConfigEdit" { New-Object System.Drawing.Size(510, 505) }
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

	# Use same theme as main form (from registry preference)
	# No need to detect here - will use Set-ThemeToForm later

	$y = 15
	$margin = 15
	$controlWidth = switch ($EditorMode) {
		"PackageList" { 380 }
		"ScriptMapping" { 470 }
		"ConfigEdit" { 470 }
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
		if ($ListName -ne "") {
			$txtListName.Enabled = $false
		}
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
		"ConfigEdit" { "Configuration Settings:" }
	}
	$editorForm.Controls.Add($lblPackages)

	$txtPackages = New-Object System.Windows.Forms.TextBox
	$txtPackages.Location = New-Object System.Drawing.Point($margin, ($y + 25))
	$txtPackages.Size = switch ($EditorMode) {
		"PackageList" { New-Object System.Drawing.Size($controlWidth, 140) }
		"ScriptMapping" { New-Object System.Drawing.Size($controlWidth, 270) }
		"ConfigEdit" { New-Object System.Drawing.Size($controlWidth, 270) }
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
	} elseif ($EditorMode -eq "ConfigEdit") {
		$listPath = Join-Path (Join-Path $Script:WorkingDir "wsb") "config.ini"
	} else {
		$listPath = if ($ListName) { Join-Path (Join-Path $Script:WorkingDir "wsb") "$ListName.txt" } else { $null }
	}

	# Variable to track original content for change detection (script scope for event handlers)
	$script:editorOriginalContent = ""

	if ($listPath -and (Test-Path $listPath)) {
		try {
			# Don't trim - preserve content including comments and empty lines
			$script:editorOriginalContent = Get-Content -Path $listPath -Raw
			# Remove trailing newline only (not all whitespace)
			if ($script:editorOriginalContent -and $script:editorOriginalContent.EndsWith("`r`n")) {
				$script:editorOriginalContent = $script:editorOriginalContent.TrimEnd("`r`n")
			}
			$txtPackages.Text = $script:editorOriginalContent
		}
		catch {
			[System.Windows.Forms.MessageBox]::Show("Error loading: $($_.Exception.Message)", "Load Error", "OK", "Error")
		}
	}

	$y += switch ($EditorMode) {
		"PackageList" { 175 }
		"ScriptMapping" { 320 }
		"ConfigEdit" { 320 }
	}

	# Help text
	$lblHelp = New-Object System.Windows.Forms.Label
	$lblHelp.Location = New-Object System.Drawing.Point($margin, $y)
	$lblHelp.Size = switch ($EditorMode) {
		"PackageList" { New-Object System.Drawing.Size($controlWidth, 50) }
		"ScriptMapping" { New-Object System.Drawing.Size($controlWidth, 70) }
		"ConfigEdit" { New-Object System.Drawing.Size($controlWidth, 70) }
	}
	$lblHelp.Text = switch ($EditorMode) {
		"PackageList" { "Example: Notepad++.Notepad++`nUse WinGet package IDs from winget search`nComments: Lines starting with # are ignored" }
		"ScriptMapping" { @"
Format: FilePattern = ScriptToExecute.ps1
Example: InstallWSB.cmd = Std-WAU.ps1
Patterns are matched against folder/file names (case-insensitive).
Wildcards: * (any characters), ? (single character)
Comments: Lines starting with # are ignored.
"@ }
		"ConfigEdit" { "INI format configuration file.`n[Lists] section: Package list states (1=enabled, 0=disabled)`n[Extensions] section: File extension mappings (ext=ListName)`nComments: Lines starting with # are ignored." }
	}
	$lblHelp.Name = 'lblHelp'  # For theme detection
	$editorForm.Controls.Add($lblHelp)

	$y += switch ($EditorMode) {
		"PackageList" { 50 }
		"ScriptMapping" { 90 }
		"ConfigEdit" { 90 }
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
		} elseif ($EditorMode -eq "ConfigEdit") {
			# ConfigEdit mode - direct to config.ini
			$wsbDir = Join-Path $Script:WorkingDir "wsb"
			if (-not (Test-Path $wsbDir)) {
				New-Item -ItemType Directory -Path $wsbDir -Force | Out-Null
			}
			$listPath = Join-Path $wsbDir "config.ini"
			$listNameValue = "sandboxtest-config"
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
			Set-Content -Path $listPath -Value $packageContent -Encoding UTF8

			# Update original content and disable Save button
			$script:editorOriginalContent = $packageContent
			$btnSave.Enabled = $false

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

	# Set initial Save button state - disabled if editing existing content
	if ($listPath -and (Test-Path $listPath)) {
		$btnSave.Enabled = $false
	}

	# Add TextChanged event to enable/disable Save button based on changes
	$txtPackages.Add_TextChanged({
		$currentContent = $txtPackages.Text.Trim()
		$hasChanged = ($currentContent -ne $script:editorOriginalContent)
		$btnSave.Enabled = $hasChanged
	})

	$editorForm.AcceptButton = $btnSave
	$editorForm.CancelButton = $btnCancel

	# Apply theme based on saved preference
	# Note: Package list editor doesn't use update button, so pass Empty color
	Set-ThemeToForm -Form $editorForm -UpdateButtonColor ([System.Drawing.Color]::Empty)

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


# Helper function to find package list by priority
function Find-PackageListByPriority {
	param(
		[Parameter(Mandatory)]
		[System.Windows.Forms.ComboBox]$ComboBox,

		[Parameter(Mandatory)]
		[string[]]$CandidateNames
	)

	foreach ($name in $CandidateNames) {
		if ($ComboBox.Items -contains $name) {
			return $name
		}
	}

	return $null
}


# Unified auto-selection function for file extensions
function Update-PackageSelectionForFileType {
	param(
		[Parameter(Mandatory)]
		[string]$FileName,

		[Parameter(Mandatory)]
		[System.Windows.Forms.CheckBox]$NetworkingCheckbox,

		[Parameter(Mandatory)]
		[System.Windows.Forms.CheckBox]$SkipWinGetCheckbox,

		[Parameter(Mandatory)]
		[System.Windows.Forms.ComboBox]$PackageComboBox,

		[Parameter(Mandatory)]
		[System.Windows.Forms.Label]$StatusLabel,

		[Parameter(Mandatory)]
		[string]$WorkingDir
	)

	# Load extension mappings from INI file
	$extensionMappings = Get-SandboxConfig -Section 'Extensions' -WorkingDir $WorkingDir

	# Get file extension (without leading dot, lowercase)
	$extension = [System.IO.Path]::GetExtension($FileName).TrimStart('.').ToLower()

	# Check if this extension has a mapping
	if (-not $extensionMappings.ContainsKey($extension)) {
		return  # No mapping for this file type
	}

	$preferredPackage = $extensionMappings[$extension]

	# Generate fallback candidate list
	$candidateNames = @($preferredPackage)

	# Add fallback: Strip "Std-" prefix if present
	if ($preferredPackage -like "Std-*") {
		$baseName = $preferredPackage -replace '^Std-', ''
		$candidateNames += $baseName
	}

	# Add alternative full names for known extensions
	$fullNameMap = @{
		'AHK' = 'AutoHotkey'
		'AU3' = 'AutoIt'
	}

	$baseName = $preferredPackage -replace '^Std-', ''
	if ($fullNameMap.ContainsKey($baseName)) {
		$fullName = $fullNameMap[$baseName]
		$candidateNames += "Std-$fullName"
		$candidateNames += $fullName
	}

	# Remove duplicates while preserving order
	$candidateNames = $candidateNames | Select-Object -Unique

	# Determine display name (strip "Std-" for status messages)
	$displayName = $preferredPackage -replace '^Std-', ''

	# Check if WinGet features are enabled
	$winGetFeaturesEnabled = $NetworkingCheckbox.Checked -and -not $SkipWinGetCheckbox.Checked

	if ($winGetFeaturesEnabled) {
		# Try to find matching package list
		$matchedPackage = Find-PackageListByPriority -ComboBox $PackageComboBox -CandidateNames $candidateNames

		if ($matchedPackage) {
			# Package list exists - auto-select it
			$PackageComboBox.SelectedItem = $matchedPackage
			$StatusLabel.Text = "Status: .$extension selected -> Auto-selected $displayName package for installation"
		} else {
			# No matching package list found
			$StatusLabel.Text = "Status: .$extension selected -> WARNING: create '$displayName.txt' in wsb\ folder!"
		}
	} elseif ($SkipWinGetCheckbox.Checked) {
		# Skip WinGet is enabled - show warning
		$StatusLabel.Text = "Status: .$extension selected -> WARNING: Uncheck 'Skip WinGet installation'!"
	} else {
		# Networking disabled - show warning
		$StatusLabel.Text = "Status: .$extension selected -> WARNING: Enable networking (WinGet)!"
	}
}


# Process selected folder or file and update form controls
function global:Update-FormFromSelection {
	param(
		[string]$SelectedPath,
		[string]$FileName = $null,
		[System.Windows.Forms.TextBox]$txtMapFolder,
		[System.Windows.Forms.TextBox]$txtSandboxFolderName,
		[System.Windows.Forms.TextBox]$txtScript,
		[System.Windows.Forms.Label]$lblStatus,
		[System.Windows.Forms.Button]$btnSaveScript,
		[System.Windows.Forms.CheckBox]$chkNetworking,
		[System.Windows.Forms.CheckBox]$chkSkipWinGet,
		[System.Windows.Forms.ComboBox]$cmbInstallPackages,
		[string]$wsbDir
	)

	# Determine if this is a file or folder selection
	$isFile = ![string]::IsNullOrWhiteSpace($FileName)

	# Get directory path
	if ($isFile) {
		$selectedDir = [System.IO.Path]::GetDirectoryName($SelectedPath)
	} else {
		$selectedDir = $SelectedPath
	}

	# Update mapped folder textbox
	$txtMapFolder.Text = $selectedDir

	# Update sandbox folder name
	if (!$isFile) {
		# Folder selected - check for WAU MSI files
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
	} else {
		# File selected - use directory name only (no WAU detection)
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

	# Handle script selection
	if ($isFile) {
		# Check if custom Std-File.ps1 exists with CUSTOM OVERRIDE header
		$stdFilePath = Join-Path $wsbDir "Std-File.ps1"
		$useCustom = $false

		if (Test-Path $stdFilePath) {
			$stdFileContent = Get-Content $stdFilePath -Raw -ErrorAction SilentlyContinue
			if ($stdFileContent -match '^\s*#\s*CUSTOM\s+OVERRIDE') {
				$useCustom = $true
			}
		}

		if ($useCustom) {
			# Custom Std-File.ps1 exists - create wrapper to call it
			# The wrapper ensures param() blocks work correctly
			$txtScript.Text = @"
`$SandboxFolderName = "$($txtSandboxFolderName.Text)"
& "`$env:USERPROFILE\Desktop\SandboxTest\Std-File.ps1" -SandboxFolderName `$SandboxFolderName -FileName "$FileName"
"@

			$lblStatus.Text = "Status: File selected -> $FileName (using CUSTOM Std-File.ps1)"

			# Set current script file so user can edit via Load... button
			$script:currentScriptFile = $stdFilePath

			# Disable Save button for wrapper (use Load... to edit actual script)
			$btnSaveScript.Enabled = $false
		} else {
			# Generate default wrapper script for Std-File.ps1
			$txtScript.Text = @"
`$SandboxFolderName = "$($txtSandboxFolderName.Text)"
& "`$env:USERPROFILE\Desktop\SandboxTest\Std-File.ps1" -SandboxFolderName `$SandboxFolderName -FileName "$FileName"
"@

			$lblStatus.Text = "Status: File selected -> $FileName (using Std-File.ps1)"

			# Disable Save button (wrapper script is auto-generated)
			$script:currentScriptFile = $null
			$btnSaveScript.Enabled = $false
		}

		# Store selected file for re-evaluation when checkboxes change
		$script:currentSelectedFile = $FileName

		# Auto-select package lists for specific file types (.py, .ahk, .au3)
		Update-PackageSelectionForFileType -FileName $FileName `
			-NetworkingCheckbox $chkNetworking `
			-SkipWinGetCheckbox $chkSkipWinGet `
			-PackageComboBox $cmbInstallPackages `
			-StatusLabel $lblStatus `
			-WorkingDir $wsbDir
	} else {
		# Folder selected - find matching script from mappings
		$matchingScript = Find-MatchingScript -Path $selectedDir
		$scriptName = $matchingScript.Replace('.ps1', '')

		# Load script using the dynamic loading function
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

			# Reset original content for default scripts (cannot be saved)
			$script:originalScriptContent = $null

			# Update Save button state
			# Check if loaded script has CUSTOM OVERRIDE header
			$hasCustomOverride = $scriptContent -match '^\s*#\s*CUSTOM\s+OVERRIDE'
			$isDefaultScript = $false
			if (-not $hasCustomOverride) {
				$isDefaultScript = Test-IsDefaultScript -FilePath $script:currentScriptFile
			}

			if ($isDefaultScript) {
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
}
# Helper function to apply theme to a dialog form
function global:Set-ThemedDialog {
	<#
	.SYNOPSIS
	Applies current theme settings to a dialog form

	.DESCRIPTION
	Internal helper function that applies the currently selected theme
	(Auto/Light/Dark/Custom) to a dialog form, including icon, colors,
	and title bar styling.

	.PARAMETER Dialog
	The Windows Forms dialog to apply theme to

	.PARAMETER ParentIcon
	Optional parent form icon to inherit
	#>
	param(
		[System.Windows.Forms.Form]$Dialog,
		[System.Drawing.Icon]$ParentIcon = $null
	)

	# Set icon
	if ($ParentIcon) {
		try {
			$Dialog.Icon = $ParentIcon
			$Dialog.ShowIcon = $true
		} catch {
			$Dialog.ShowIcon = $false
		}
	} else {
		$Dialog.ShowIcon = $false
	}

	# Get current theme and apply
	$currentTheme = Get-SandboxStartThemePreference
	$useDarkTitleBar = $false

	if ($currentTheme -eq "Dark") {
		Set-DarkModeTheme -Control $Dialog
		$useDarkTitleBar = $true
	} elseif ($currentTheme -eq "Light") {
		Set-LightModeTheme -Control $Dialog
		$useDarkTitleBar = $false
	} elseif ($currentTheme -eq "Custom") {
		$customColors = Get-SandboxStartCustomColors
		Set-CustomTheme -Control $Dialog -CustomColors $customColors
		$bgRgb = $customColors.BackColor -split ','
		$bgColor = [System.Drawing.Color]::FromArgb([int]$bgRgb[0], [int]$bgRgb[1], [int]$bgRgb[2])
		$useDarkTitleBar = Test-ColorIsDark -Color $bgColor
	} elseif ($currentTheme -eq "Auto") {
		if (Test-SystemUsesLightTheme) {
			Set-LightModeTheme -Control $Dialog
			$useDarkTitleBar = $false
		} else {
			Set-DarkModeTheme -Control $Dialog
			$useDarkTitleBar = $true
		}
	}

	# Apply title bar theme
	Set-DarkTitleBar -Form $Dialog -UseDarkMode $useDarkTitleBar
}

# Helper function to show themed input dialog
function global:Show-ThemedInputDialog {
	<#
	.SYNOPSIS
	Displays a themed input dialog for text entry

	.DESCRIPTION
	Shows a custom dialog with theming applied that prompts the user
	for text input. Supports all theme types (Auto/Light/Dark/Custom)
	and provides OK/Cancel buttons.

	.PARAMETER Title
	The dialog window title

	.PARAMETER Prompt
	The prompt text displayed above the input field

	.PARAMETER DefaultValue
	Optional default value pre-populated in the textbox

	.PARAMETER ParentIcon
	Optional parent form icon to inherit

	.RETURNS
	String containing user input, or $null if canceled
	#>
	param(
		[string]$Title = "Input",
		[string]$Prompt = "Enter value:",
		[string]$DefaultValue = "",
		[System.Drawing.Icon]$ParentIcon = $null
	)

	# Create dialog
	$inputDialog = New-Object System.Windows.Forms.Form
	$inputDialog.Text = $Title
	$inputDialog.Size = New-Object System.Drawing.Size(400, 160)
	$inputDialog.StartPosition = "CenterParent"
	$inputDialog.FormBorderStyle = "FixedDialog"
	$inputDialog.MaximizeBox = $false
	$inputDialog.MinimizeBox = $false

	# Label
	$inputLabel = New-Object System.Windows.Forms.Label
	$inputLabel.Location = New-Object System.Drawing.Point(10, 20)
	$inputLabel.Size = New-Object System.Drawing.Size(360, 20)
	$inputLabel.Text = $Prompt
	$inputDialog.Controls.Add($inputLabel)

	# TextBox
	$inputTextBox = New-Object System.Windows.Forms.TextBox
	$inputTextBox.Location = New-Object System.Drawing.Point(10, 45)
	$inputTextBox.Size = New-Object System.Drawing.Size(360, 20)
	$inputTextBox.Text = $DefaultValue
	$inputDialog.Controls.Add($inputTextBox)

	# OK Button
	$inputOkButton = New-Object System.Windows.Forms.Button
	$inputOkButton.Location = New-Object System.Drawing.Point(190, 80)
	$inputOkButton.Size = New-Object System.Drawing.Size(85, 25)
	$inputOkButton.Text = "OK"
	$inputOkButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$inputDialog.AcceptButton = $inputOkButton
	$inputDialog.Controls.Add($inputOkButton)

	# Cancel Button
	$inputCancelButton = New-Object System.Windows.Forms.Button
	$inputCancelButton.Location = New-Object System.Drawing.Point(285, 80)
	$inputCancelButton.Size = New-Object System.Drawing.Size(85, 25)
	$inputCancelButton.Text = "Cancel"
	$inputCancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$inputDialog.CancelButton = $inputCancelButton
	$inputDialog.Controls.Add($inputCancelButton)

	# Apply theme
	Set-ThemedDialog -Dialog $inputDialog -ParentIcon $ParentIcon | Out-Null

	# Show dialog and capture result
	$result = $inputDialog.ShowDialog()

	# Capture text value before disposing
	$returnValue = $null
	if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
		$text = $inputTextBox.Text
		if (-not [string]::IsNullOrWhiteSpace($text)) {
			$returnValue = [string]$text
		}
	}

	# Dispose dialog
	$inputDialog.Dispose() | Out-Null

	# Return text value or null
	Write-Output $returnValue
}

# Helper function to show themed message dialog
function global:Show-ThemedMessageDialog {
	<#
	.SYNOPSIS
	Displays a themed message dialog

	.DESCRIPTION
	Shows a custom message dialog with theming applied. Supports various
	button configurations (OK, OKCancel, YesNo, YesNoCancel) and icon
	types (Information, Warning, Error, Question).

	.PARAMETER Title
	The dialog window title

	.PARAMETER Message
	The message text to display (supports multi-line with `n)

	.PARAMETER Buttons
	Button configuration: "OK", "OKCancel", "YesNo", "YesNoCancel"

	.PARAMETER Icon
	Icon type: "Information", "Warning", "Error", "Question"

	.PARAMETER ParentIcon
	Optional parent form icon to inherit

	.RETURNS
	System.Windows.Forms.DialogResult indicating which button was clicked
	#>
	param(
		[string]$Title = "Message",
		[string]$Message = "",
		[string]$Buttons = "OK",
		[string]$Icon = "Information",
		[System.Drawing.Icon]$ParentIcon = $null
	)

	# Calculate dialog height based on message length
	$messageLines = ($Message -split "`n").Count
	$messageHeight = [Math]::Max(60, $messageLines * 20 + 20)
	$dialogHeight = $messageHeight + 100

	# Create dialog
	$messageDialog = New-Object System.Windows.Forms.Form
	$messageDialog.Text = $Title
	$messageDialog.Size = New-Object System.Drawing.Size(450, $dialogHeight)
	$messageDialog.StartPosition = "CenterParent"
	$messageDialog.FormBorderStyle = "FixedDialog"
	$messageDialog.MaximizeBox = $false
	$messageDialog.MinimizeBox = $false

	# Message label
	$messageLabel = New-Object System.Windows.Forms.Label
	$messageLabel.Location = New-Object System.Drawing.Point(10, 20)
	$messageLabel.Size = New-Object System.Drawing.Size(420, $messageHeight)
	$messageLabel.Text = $Message
	$messageDialog.Controls.Add($messageLabel)

	# Calculate button Y position
	$buttonY = $messageHeight + 30

	# Create buttons based on button type
	if ($Buttons -eq "OK") {
		$okButton = New-Object System.Windows.Forms.Button
		$okButton.Location = New-Object System.Drawing.Point(175, $buttonY)
		$okButton.Size = New-Object System.Drawing.Size(85, 25)
		$okButton.Text = "OK"
		$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
		$messageDialog.AcceptButton = $okButton
		$messageDialog.Controls.Add($okButton)
	}
	elseif ($Buttons -eq "OKCancel") {
		$okButton = New-Object System.Windows.Forms.Button
		$okButton.Location = New-Object System.Drawing.Point(240, $buttonY)
		$okButton.Size = New-Object System.Drawing.Size(85, 25)
		$okButton.Text = "OK"
		$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
		$messageDialog.AcceptButton = $okButton
		$messageDialog.Controls.Add($okButton)

		$cancelButton = New-Object System.Windows.Forms.Button
		$cancelButton.Location = New-Object System.Drawing.Point(335, $buttonY)
		$cancelButton.Size = New-Object System.Drawing.Size(85, 25)
		$cancelButton.Text = "Cancel"
		$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
		$messageDialog.CancelButton = $cancelButton
		$messageDialog.Controls.Add($cancelButton)
	}
	elseif ($Buttons -eq "YesNo") {
		$yesButton = New-Object System.Windows.Forms.Button
		$yesButton.Location = New-Object System.Drawing.Point(240, $buttonY)
		$yesButton.Size = New-Object System.Drawing.Size(85, 25)
		$yesButton.Text = "Yes"
		$yesButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
		$messageDialog.AcceptButton = $yesButton
		$messageDialog.Controls.Add($yesButton)

		$noButton = New-Object System.Windows.Forms.Button
		$noButton.Location = New-Object System.Drawing.Point(335, $buttonY)
		$noButton.Size = New-Object System.Drawing.Size(85, 25)
		$noButton.Text = "No"
		$noButton.DialogResult = [System.Windows.Forms.DialogResult]::No
		$messageDialog.CancelButton = $noButton
		$messageDialog.Controls.Add($noButton)
	}
	elseif ($Buttons -eq "YesNoCancel") {
		$yesButton = New-Object System.Windows.Forms.Button
		$yesButton.Location = New-Object System.Drawing.Point(145, $buttonY)
		$yesButton.Size = New-Object System.Drawing.Size(85, 25)
		$yesButton.Text = "Yes"
		$yesButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
		$messageDialog.Controls.Add($yesButton)

		$noButton = New-Object System.Windows.Forms.Button
		$noButton.Location = New-Object System.Drawing.Point(240, $buttonY)
		$noButton.Size = New-Object System.Drawing.Size(85, 25)
		$noButton.Text = "No"
		$noButton.DialogResult = [System.Windows.Forms.DialogResult]::No
		$messageDialog.Controls.Add($noButton)

		$cancelButton = New-Object System.Windows.Forms.Button
		$cancelButton.Location = New-Object System.Drawing.Point(335, $buttonY)
		$cancelButton.Size = New-Object System.Drawing.Size(85, 25)
		$cancelButton.Text = "Cancel"
		$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
		$messageDialog.CancelButton = $cancelButton
		$messageDialog.Controls.Add($cancelButton)
	}

	# Apply theme
	Set-ThemedDialog -Dialog $messageDialog -ParentIcon $ParentIcon

	# Show dialog and return result
	return $messageDialog.ShowDialog()
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

function global:Set-DarkModeTheme {
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
		# Skip color picker swatch panels and preview panel (they need to keep their custom colors)
		if ($child.Tag -eq "ColorPickerSwatch" -or $child.Tag -eq "ColorPreviewPanel") {
			continue
		}
		Set-DarkModeTheme -Control $child -UpdateButtonBackColor $UpdateButtonBackColor
	}
}

function global:Set-LightModeTheme {
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
		# Skip color picker swatch panels and preview panel (they need to keep their custom colors)
		if ($child.Tag -eq "ColorPickerSwatch" -or $child.Tag -eq "ColorPreviewPanel") {
			continue
		}
		Set-LightModeTheme -Control $child -UpdateButtonBackColor $UpdateButtonBackColor
	}
}

function global:Set-DarkTitleBar {
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

	# Apply title bar theme immediately if form is already created, otherwise use Shown event
	if ($Form.Handle -ne [IntPtr]::Zero) {
		# Form already has a window handle - apply immediately
		try {
			# DWMWA_USE_IMMERSIVE_DARK_MODE = 20
			$titleBarMode = if ($UseDarkMode) { 1 } else { 0 }
			[DarkMode.DwmApi]::DwmSetWindowAttribute($Form.Handle, 20, [ref]$titleBarMode, 4) | Out-Null
		}
		catch {
			# Silent fail - not critical if title bar theming doesn't work
			Write-Verbose "Failed to set title bar theme: $($_.Exception.Message)"
		}
	}
	else {
		# Form not yet shown - use Shown event to apply when window handle is created
		$darkMode = $UseDarkMode
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
}

function global:Set-CustomTheme {
	<#
	.SYNOPSIS
	Applies custom user-defined colors to a Windows Form and all its controls recursively

	.PARAMETER Control
	The form or control to apply custom theme to

	.PARAMETER CustomColors
	Hashtable containing 6 color elements as RGB strings ("R,G,B")

	.PARAMETER UpdateButtonBackColor
	Optional. BackColor override for the update button (adaptive green)
	#>

	param(
		[Parameter(Mandatory)]
		[System.Windows.Forms.Control]$Control,

		[Parameter(Mandatory)]
		[hashtable]$CustomColors,

		[Parameter(Mandatory = $false)]
		[System.Drawing.Color]$UpdateButtonBackColor = [System.Drawing.Color]::Empty
	)

	# Parse color strings from hashtable to Color objects
	try {
		$backColor = [System.Drawing.Color]::FromArgb([int]::Parse($CustomColors.BackColor.Split(',')[0]), [int]::Parse($CustomColors.BackColor.Split(',')[1]), [int]::Parse($CustomColors.BackColor.Split(',')[2]))
		$foreColor = [System.Drawing.Color]::FromArgb([int]::Parse($CustomColors.ForeColor.Split(',')[0]), [int]::Parse($CustomColors.ForeColor.Split(',')[1]), [int]::Parse($CustomColors.ForeColor.Split(',')[2]))
		$buttonBackColor = [System.Drawing.Color]::FromArgb([int]::Parse($CustomColors.ButtonBackColor.Split(',')[0]), [int]::Parse($CustomColors.ButtonBackColor.Split(',')[1]), [int]::Parse($CustomColors.ButtonBackColor.Split(',')[2]))
		$textBoxBackColor = [System.Drawing.Color]::FromArgb([int]::Parse($CustomColors.TextBoxBackColor.Split(',')[0]), [int]::Parse($CustomColors.TextBoxBackColor.Split(',')[1]), [int]::Parse($CustomColors.TextBoxBackColor.Split(',')[2]))
		$grayLabelColor = [System.Drawing.Color]::FromArgb([int]::Parse($CustomColors.GrayLabelColor.Split(',')[0]), [int]::Parse($CustomColors.GrayLabelColor.Split(',')[1]), [int]::Parse($CustomColors.GrayLabelColor.Split(',')[2]))
		$updateBtnColor = [System.Drawing.Color]::FromArgb([int]::Parse($CustomColors.UpdateButtonColor.Split(',')[0]), [int]::Parse($CustomColors.UpdateButtonColor.Split(',')[1]), [int]::Parse($CustomColors.UpdateButtonColor.Split(',')[2]))
	}
	catch {
		Write-Warning "Failed to parse custom colors, falling back to dark mode defaults: $($_.Exception.Message)"
		# Fallback to dark mode if parsing fails
		$backColor = [System.Drawing.Color]::FromArgb(32, 32, 32)
		$foreColor = [System.Drawing.Color]::White
		$buttonBackColor = [System.Drawing.Color]::FromArgb(70, 70, 70)
		$textBoxBackColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
		$grayLabelColor = [System.Drawing.Color]::FromArgb(180, 180, 180)
		$updateBtnColor = [System.Drawing.Color]::FromArgb(60, 120, 60)
	}

	# Apply base colors
	$Control.BackColor = $backColor
	$Control.ForeColor = $foreColor

	# Special handling by control type (similar to Set-DarkModeTheme logic)
	if ($Control -is [System.Windows.Forms.Button]) {
		$Control.BackColor = $buttonBackColor

		# Special case: Update button
		if ($Control.Name -eq 'btnUpdate' -or ($Control.Text -eq [char]0x2B06)) {
			if ($UpdateButtonBackColor -ne [System.Drawing.Color]::Empty) {
				$Control.BackColor = $UpdateButtonBackColor
			}
			else {
				$Control.BackColor = $updateBtnColor
			}
		}
	}
	elseif ($Control -is [System.Windows.Forms.TextBox]) {
		$Control.BackColor = $textBoxBackColor
		$Control.ForeColor = $foreColor
	}
	elseif ($Control -is [System.Windows.Forms.ComboBox]) {
		$Control.BackColor = $textBoxBackColor
		$Control.ForeColor = $foreColor
	}
	elseif ($Control -is [System.Windows.Forms.CheckBox]) {
		$Control.BackColor = $backColor
		$Control.ForeColor = $foreColor
	}
	elseif ($Control -is [System.Windows.Forms.Label]) {
		# Check if this is a gray label (help text)
		if ($Control.Name -eq 'lblHelp' -or $Control.ForeColor.ToArgb() -eq [System.Drawing.Color]::Gray.ToArgb()) {
			$Control.ForeColor = $grayLabelColor
		}
		else {
			$Control.ForeColor = $foreColor
		}
	}

	# Recursively apply to child controls
	foreach ($child in $Control.Controls) {
		# Skip color picker swatch panels and preview panel (they need to keep their custom colors)
		if ($child.Tag -eq "ColorPickerSwatch" -or $child.Tag -eq "ColorPreviewPanel") {
			continue
		}
		Set-CustomTheme -Control $child -CustomColors $CustomColors -UpdateButtonBackColor $UpdateButtonBackColor
	}
}

function global:Set-ThemeToForm {
	<#
	.SYNOPSIS
	Applies the selected theme to a Windows Form based on user preference

	.PARAMETER Form
	The form to apply theme to

	.PARAMETER UpdateButtonColor
	The color to use for the update button
	#>

	param(
		[Parameter(Mandatory)]
		[System.Windows.Forms.Form]$Form,

		[Parameter(Mandatory)]
		[System.Drawing.Color]$UpdateButtonColor
	)

	$themeMode = Get-SandboxStartThemePreference

	switch ($themeMode) {
		"Auto" {
			$useDarkMode = Get-WindowsThemeSetting
			if ($useDarkMode) {
				Set-DarkModeTheme -Control $Form -UpdateButtonBackColor $UpdateButtonColor
				Set-DarkTitleBar -Form $Form -UseDarkMode $true
			}
			else {
				Set-LightModeTheme -Control $Form -UpdateButtonBackColor $UpdateButtonColor
				Set-DarkTitleBar -Form $Form -UseDarkMode $false
			}
		}
		"Light" {
			Set-LightModeTheme -Control $Form -UpdateButtonBackColor $UpdateButtonColor
			Set-DarkTitleBar -Form $Form -UseDarkMode $false
		}
		"Dark" {
			Set-DarkModeTheme -Control $Form -UpdateButtonBackColor $UpdateButtonColor
			Set-DarkTitleBar -Form $Form -UseDarkMode $true
		}
		"Custom" {
			$customColors = Get-SandboxStartCustomColors
			Set-CustomTheme -Control $Form -CustomColors $customColors -UpdateButtonBackColor $UpdateButtonColor
			# Detect if background is dark or light for title bar
			try {
				$bgColorParts = $customColors.BackColor.Split(',')
				$bgColor = [System.Drawing.Color]::FromArgb([int]::Parse($bgColorParts[0]), [int]::Parse($bgColorParts[1]), [int]::Parse($bgColorParts[2]))
				$isDark = Test-ColorIsDark -Color $bgColor
				Set-DarkTitleBar -Form $Form -UseDarkMode $isDark
			}
			catch {
				# Fallback to dark if color parsing fails
				Set-DarkTitleBar -Form $Form -UseDarkMode $true
			}
		}
	}

	$Form.Refresh()
}

function global:Show-ThemeContextMenu {
	<#
	.SYNOPSIS
	Creates a context menu for theme selection

	.PARAMETER Form
	The form to attach the menu to

	.PARAMETER UpdateButtonColor
	The color to use for the update button

	.PARAMETER WorkingDir
	Optional working directory for shell integration (SandboxStart only)

	.PARAMETER TestRegKey
	Optional scriptblock for testing registry keys

	.PARAMETER TestContextMenu
	Optional scriptblock for testing context menu installation

	.PARAMETER UpdateContextMenu
	Optional scriptblock for updating context menu

	.PARAMETER AppIcon
	Optional icon for message dialogs

	.OUTPUTS
	ContextMenuStrip object
	#>

	param(
		[Parameter(Mandatory)]
		[System.Windows.Forms.Form]$Form,

		[Parameter(Mandatory)]
		[System.Drawing.Color]$UpdateButtonColor,

		[string]$WorkingDir,

		[scriptblock]$TestRegKey,

		[scriptblock]$TestContextMenu,

		[scriptblock]$UpdateContextMenu,

		[System.Drawing.Icon]$AppIcon
	)

	# Get current theme preference
	$currentTheme = Get-SandboxStartThemePreference

	# Create context menu
	$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip

	# Theme context menu to match form
	$menuBackColor = $Form.BackColor
	$menuForeColor = $Form.ForeColor
	$contextMenu.BackColor = $menuBackColor
	$contextMenu.ForeColor = $menuForeColor

	# Store references for use in event handlers
	$menuForm = $Form
	$menuColor = $UpdateButtonColor
	$menuWorkingDir = $WorkingDir
	$menuTestRegKey = $TestRegKey
	$menuTestContextMenu = $TestContextMenu
	$menuUpdateContextMenu = $UpdateContextMenu
	$menuAppIcon = $AppIcon

	# Add "Theme" header (disabled, acts as label)
	$headerItem = New-Object System.Windows.Forms.ToolStripMenuItem
	$headerItem.Text = "Theme"
	$headerItem.Enabled = $false
	$headerItem.BackColor = $menuBackColor
	$headerItem.ForeColor = $menuForeColor
	$contextMenu.Items.Add($headerItem) | Out-Null

	# Add separator
	$contextMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null

	# Add theme options
	$autoItem = New-Object System.Windows.Forms.ToolStripMenuItem
	$autoItem.Text = "Auto (Follow System)"
	$autoItem.Checked = ($currentTheme -eq "Auto")
	$autoItem.BackColor = $menuBackColor
	$autoItem.ForeColor = $menuForeColor
	$autoItem.Add_Click({
			Set-SandboxStartThemePreference -ThemeMode "Auto"
			Set-ThemeToForm -Form $menuForm -UpdateButtonColor $menuColor
			$menuForm.ContextMenuStrip = Show-ThemeContextMenu -Form $menuForm -UpdateButtonColor $menuColor -WorkingDir $menuWorkingDir -TestRegKey $menuTestRegKey -TestContextMenu $menuTestContextMenu -UpdateContextMenu $menuUpdateContextMenu -AppIcon $menuAppIcon
		}.GetNewClosure())
	$contextMenu.Items.Add($autoItem) | Out-Null

	$lightItem = New-Object System.Windows.Forms.ToolStripMenuItem
	$lightItem.Text = "Light"
	$lightItem.Checked = ($currentTheme -eq "Light")
	$lightItem.BackColor = $menuBackColor
	$lightItem.ForeColor = $menuForeColor
	$lightItem.Add_Click({
			Set-SandboxStartThemePreference -ThemeMode "Light"
			Set-ThemeToForm -Form $menuForm -UpdateButtonColor $menuColor
			$menuForm.ContextMenuStrip = Show-ThemeContextMenu -Form $menuForm -UpdateButtonColor $menuColor -WorkingDir $menuWorkingDir -TestRegKey $menuTestRegKey -TestContextMenu $menuTestContextMenu -UpdateContextMenu $menuUpdateContextMenu -AppIcon $menuAppIcon
		}.GetNewClosure())
	$contextMenu.Items.Add($lightItem) | Out-Null

	$darkItem = New-Object System.Windows.Forms.ToolStripMenuItem
	$darkItem.Text = "Dark"
	$darkItem.Checked = ($currentTheme -eq "Dark")
	$darkItem.BackColor = $menuBackColor
	$darkItem.ForeColor = $menuForeColor
	$darkItem.Add_Click({
			Set-SandboxStartThemePreference -ThemeMode "Dark"
			Set-ThemeToForm -Form $menuForm -UpdateButtonColor $menuColor
			$menuForm.ContextMenuStrip = Show-ThemeContextMenu -Form $menuForm -UpdateButtonColor $menuColor -WorkingDir $menuWorkingDir -TestRegKey $menuTestRegKey -TestContextMenu $menuTestContextMenu -UpdateContextMenu $menuUpdateContextMenu -AppIcon $menuAppIcon
		}.GetNewClosure())
	$contextMenu.Items.Add($darkItem) | Out-Null

	# Add separator
	$contextMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null

	# Add custom colors option
	$customItem = New-Object System.Windows.Forms.ToolStripMenuItem
	$customItem.Text = "Custom..."
	$customItem.Checked = ($currentTheme -eq "Custom")
	$customItem.BackColor = $menuBackColor
	$customItem.ForeColor = $menuForeColor
	$customItem.Add_Click({
			Show-ColorPickerDialog -ParentForm $menuForm -UpdateButtonColor $menuColor -WorkingDir $menuWorkingDir -TestRegKey $menuTestRegKey -TestContextMenu $menuTestContextMenu -UpdateContextMenu $menuUpdateContextMenu -AppIcon $menuAppIcon
		}.GetNewClosure())
	$contextMenu.Items.Add($customItem) | Out-Null

	# Add Context Menu Integration if WorkingDir is provided and SandboxStart.ps1 exists
	if ($WorkingDir -and $TestContextMenu -and $UpdateContextMenu) {
		$sandboxStartScript = Join-Path $WorkingDir 'SandboxStart.ps1'
		if (Test-Path $sandboxStartScript) {
			# Create menu item
			$menuContextIntegration = New-Object System.Windows.Forms.ToolStripMenuItem
			$menuContextIntegration.BackColor = $menuBackColor
			$menuContextIntegration.ForeColor = $menuForeColor

			$isInstalled = & $TestContextMenu

			$menuContextIntegration.Text = if ($isInstalled) {
				"Disable Context Menu Integration"
			} else {
				"Enable Context Menu Integration"
			}

			$menuContextIntegration.Add_Click({
				param($menuItem, $e)
				$currentlyInstalled = & $menuTestContextMenu

				if ($currentlyInstalled) {
					# Disable
					$result = Show-ThemedMessageDialog `
						-Title "Disable Context Menu Integration" `
						-Message "Remove 'Test in Windows Sandbox' from folder and file context menus?" `
						-Buttons "OKCancel" `
						-Icon "Question" `
						-ParentIcon $menuAppIcon

					if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
						if (& $menuUpdateContextMenu -WorkingDir $menuWorkingDir -Remove $true) {
							Show-ThemedMessageDialog `
								-Title "Success" `
								-Message "Context menu integration has been disabled." `
								-Buttons "OK" `
								-Icon "Information" `
								-ParentIcon $menuAppIcon
							$menuItem.Text = "Enable Context Menu Integration"
						}
					}
				}
				else {
					# Enable
					$result = Show-ThemedMessageDialog `
						-Title "Enable Context Menu Integration" `
						-Message "Add 'Test in Windows Sandbox' to folder and file context menus?`n`nThis allows you to right-click folders/files and test them directly." `
						-Buttons "OKCancel" `
						-Icon "Question" `
						-ParentIcon $menuAppIcon

					if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
						if (& $menuUpdateContextMenu -WorkingDir $menuWorkingDir -Remove $false) {
							Show-ThemedMessageDialog `
								-Title "Success" `
								-Message "Context menu integration has been enabled.`n`nYou can now right-click folders and files to test them in Windows Sandbox." `
								-Buttons "OK" `
								-Icon "Information" `
								-ParentIcon $menuAppIcon
							$menuItem.Text = "Disable Context Menu Integration"
						}
					}
				}
			}.GetNewClosure())

			# Insert at the beginning of the context menu (before theme options)
			$contextMenu.Items.Insert(0, (New-Object System.Windows.Forms.ToolStripSeparator))
			$contextMenu.Items.Insert(0, $menuContextIntegration)
		}
	}

	return $contextMenu
}

function global:Show-ColorPickerDialog {
	<#
	.SYNOPSIS
	Shows advanced color picker dialog for customizing theme colors

	.PARAMETER ParentForm
	The parent form

	.PARAMETER UpdateButtonColor
	The color to use for the update button

	.PARAMETER WorkingDir
	Optional working directory for shell integration

	.PARAMETER TestRegKey
	Optional scriptblock for testing registry keys

	.PARAMETER TestContextMenu
	Optional scriptblock for testing context menu installation

	.PARAMETER UpdateContextMenu
	Optional scriptblock for updating context menu

	.PARAMETER AppIcon
	Optional icon for message dialogs
	#>

	param(
		[Parameter(Mandatory)]
		[System.Windows.Forms.Form]$ParentForm,

		[Parameter(Mandatory)]
		[System.Drawing.Color]$UpdateButtonColor,

		[string]$WorkingDir,

		[scriptblock]$TestRegKey,

		[scriptblock]$TestContextMenu,

		[scriptblock]$UpdateContextMenu,

		[System.Drawing.Icon]$AppIcon
	)

	# Save original theme state (for Cancel button to restore)
	$originalTheme = Get-SandboxStartThemePreference
	$originalColors = Get-SandboxStartCustomColors

	# Store shell integration params for closures
	$localWorkingDir = $WorkingDir
	$localTestRegKey = $TestRegKey
	$localTestContextMenu = $TestContextMenu
	$localUpdateContextMenu = $UpdateContextMenu
	$localAppIcon = $AppIcon

	# Load current custom colors
	$currentColors = Get-SandboxStartCustomColors

	# Define preset color schemes
	$presets = @{
		"Visual Studio Dark (Default)" = @{
			BackColor = "45,45,48"
			ForeColor = "241,241,241"
			ButtonBackColor = "62,62,66"
			TextBoxBackColor = "37,37,38"
			GrayLabelColor = "133,133,133"
			UpdateButtonColor = "0,122,204"
		}
		"Visual Studio Light" = @{
			BackColor = "246,246,246"
			ForeColor = "30,30,30"
			ButtonBackColor = "238,238,238"
			TextBoxBackColor = "255,255,255"
			GrayLabelColor = "120,120,120"
			UpdateButtonColor = "0,122,204"
		}
		"Solarized Light" = @{
			BackColor = "253,246,227"
			ForeColor = "101,123,131"
			ButtonBackColor = "238,232,213"
			TextBoxBackColor = "253,246,227"
			GrayLabelColor = "147,161,161"
			UpdateButtonColor = "181,137,0"
		}
		"GitHub Light" = @{
			BackColor = "255,255,255"
			ForeColor = "36,41,46"
			ButtonBackColor = "250,251,252"
			TextBoxBackColor = "255,255,255"
			GrayLabelColor = "106,115,125"
			UpdateButtonColor = "3,102,214"
		}
		"GitHub Dark" = @{
			BackColor = "13,17,23"
			ForeColor = "230,237,243"
			ButtonBackColor = "33,38,45"
			TextBoxBackColor = "22,27,34"
			GrayLabelColor = "139,148,158"
			UpdateButtonColor = "88,166,255"
		}
		"Monokai" = @{
			BackColor = "39,40,34"
			ForeColor = "248,248,242"
			ButtonBackColor = "73,72,62"
			TextBoxBackColor = "30,31,27"
			GrayLabelColor = "117,113,94"
			UpdateButtonColor = "102,217,239"
		}
		"Dracula" = @{
			BackColor = "40,42,54"
			ForeColor = "248,248,242"
			ButtonBackColor = "68,71,90"
			TextBoxBackColor = "33,34,44"
			GrayLabelColor = "98,114,164"
			UpdateButtonColor = "139,233,253"
		}
		"Nord" = @{
			BackColor = "46,52,64"
			ForeColor = "236,239,244"
			ButtonBackColor = "59,66,82"
			TextBoxBackColor = "39,44,55"
			GrayLabelColor = "129,161,193"
			UpdateButtonColor = "136,192,208"
		}
		"Matrix Green" = @{
			BackColor = "2,2,4"
			ForeColor = "0,255,65"
			ButtonBackColor = "32,72,41"
			TextBoxBackColor = "13,2,8"
			GrayLabelColor = "128,206,135"
			UpdateButtonColor = "34,180,85"
		}
		"Hacker Terminal" = @{
			BackColor = "34,34,34"
			ForeColor = "183,206,66"
			ButtonBackColor = "50,134,76"
			TextBoxBackColor = "40,40,40"
			GrayLabelColor = "100,149,104"
			UpdateButtonColor = "189,224,119"
		}
		"Cyberpunk 2077" = @{
			BackColor = "58,10,77"
			ForeColor = "249,197,78"
			ButtonBackColor = "164,45,180"
			TextBoxBackColor = "46,8,61"
			GrayLabelColor = "255,92,138"
			UpdateButtonColor = "249,197,78"
		}
	}

	# Create dialog form
	$dialog = New-Object System.Windows.Forms.Form
	$dialog.Text = "Custom Theme Colors"
	$dialog.Size = New-Object System.Drawing.Size(520, 670)
	$dialog.StartPosition = "CenterParent"
	$dialog.FormBorderStyle = "FixedDialog"
	$dialog.MaximizeBox = $false
	$dialog.MinimizeBox = $false

	# Set form icon (same as main form)
	try {
		if ($Script:AppIcon) {
			$dialog.Icon = $Script:AppIcon
			$dialog.ShowIcon = $true
		}
		else {
			$dialog.ShowIcon = $false
		}
	}
	catch {
		$dialog.ShowIcon = $false
	}

	# Current Y position for controls
	$yPos = 20

	# Preset scheme label
	$presetLabel = New-Object System.Windows.Forms.Label
	$presetLabel.Location = New-Object System.Drawing.Point(20, $yPos)
	$presetLabel.Size = New-Object System.Drawing.Size(100, 20)
	$presetLabel.Text = "Preset Schemes:"
	$dialog.Controls.Add($presetLabel)

	# Preset dropdown
	$presetCombo = New-Object System.Windows.Forms.ComboBox
	$comboY = $yPos - 2
	$presetCombo.Location = New-Object System.Drawing.Point(120, $comboY)
	$presetCombo.Size = New-Object System.Drawing.Size(360, 25)
	$presetCombo.DropDownStyle = "DropDownList"
	$presetCombo.Items.Add("Current Colors") | Out-Null
	foreach ($presetName in $presets.Keys | Sort-Object) {
		$presetCombo.Items.Add($presetName) | Out-Null
	}
	$presetCombo.SelectedIndex = 0
	$dialog.Controls.Add($presetCombo)

	# Define themes directory for Save/Load Theme buttons
	$themesDir = Join-Path $Script:WorkingDir "themes"

	# Save Theme button (right-aligned)
	$btnSaveTheme = New-Object System.Windows.Forms.Button
	$btnSaveTheme.Location = New-Object System.Drawing.Point(250, ($yPos + 30))
	$btnSaveTheme.Size = New-Object System.Drawing.Size(110, 25)
	$btnSaveTheme.Text = "Save Theme..."
	$dialog.Controls.Add($btnSaveTheme)

	# Load Theme button (right-aligned, next to Save)
	$btnLoadTheme = New-Object System.Windows.Forms.Button
	$btnLoadTheme.Location = New-Object System.Drawing.Point(370, ($yPos + 30))
	$btnLoadTheme.Size = New-Object System.Drawing.Size(110, 25)
	$btnLoadTheme.Text = "Load Theme..."
	$dialog.Controls.Add($btnLoadTheme)

	$yPos = $yPos + 70

	# Group box for color elements
	$colorGroup = New-Object System.Windows.Forms.GroupBox
	$colorGroup.Location = New-Object System.Drawing.Point(20, $yPos)
	$colorGroup.Size = New-Object System.Drawing.Size(460, 240)
	$colorGroup.Text = "Color Elements"
	$dialog.Controls.Add($colorGroup)

	# Helper function to create color picker row
	$colorPickerRows = @{}
	$rowYPos = 25

	# Create 6 color picker rows manually (inline to avoid scope issues)
	# Background Color
	$rgb = $currentColors.BackColor -split ','
	$initialColor = [System.Drawing.Color]::FromArgb([int]$rgb[0], [int]$rgb[1], [int]$rgb[2])

	$labelY = $rowYPos + 5
	$label = New-Object System.Windows.Forms.Label
	$label.Location = New-Object System.Drawing.Point(10, $labelY)
	$label.Size = New-Object System.Drawing.Size(160, 20)
	$label.Text = "Background Color:"
	$colorGroup.Controls.Add($label)

	$panel = New-Object System.Windows.Forms.Panel
	$panel.Location = New-Object System.Drawing.Point(175, $rowYPos)
	$panel.Size = New-Object System.Drawing.Size(40, 25)
	$panel.BorderStyle = "Fixed3D"
	$panel.BackColor = $initialColor
	$panel.Tag = "ColorPickerSwatch"
	$colorGroup.Controls.Add($panel)

	$rgbLabel = New-Object System.Windows.Forms.Label
	$rgbLabel.Location = New-Object System.Drawing.Point(220, $labelY)
	$rgbLabel.Size = New-Object System.Drawing.Size(100, 20)
	$rgbLabel.Text = "RGB: $($currentColors.BackColor)"
	$colorGroup.Controls.Add($rgbLabel)

	$button = New-Object System.Windows.Forms.Button
	$button.Location = New-Object System.Drawing.Point(330, $rowYPos)
	$button.Size = New-Object System.Drawing.Size(110, 25)
	$button.Text = "Choose..."
	$button.Tag = "BackColor"
	$colorGroup.Controls.Add($button)

	$colorPickerRows["BackColor"] = @{Panel = $panel; RgbLabel = $rgbLabel; Button = $button; ColorKey = "BackColor"}
	$rowYPos = $rowYPos + 35

	# Text Color
	$rgb = $currentColors.ForeColor -split ','
	$initialColor = [System.Drawing.Color]::FromArgb([int]$rgb[0], [int]$rgb[1], [int]$rgb[2])

	$labelY = $rowYPos + 5
	$label = New-Object System.Windows.Forms.Label
	$label.Location = New-Object System.Drawing.Point(10, $labelY)
	$label.Size = New-Object System.Drawing.Size(160, 20)
	$label.Text = "Text Color:"
	$colorGroup.Controls.Add($label)

	$panel = New-Object System.Windows.Forms.Panel
	$panel.Location = New-Object System.Drawing.Point(175, $rowYPos)
	$panel.Size = New-Object System.Drawing.Size(40, 25)
	$panel.BorderStyle = "Fixed3D"
	$panel.BackColor = $initialColor
	$panel.Tag = "ColorPickerSwatch"
	$colorGroup.Controls.Add($panel)

	$rgbLabel = New-Object System.Windows.Forms.Label
	$rgbLabel.Location = New-Object System.Drawing.Point(220, $labelY)
	$rgbLabel.Size = New-Object System.Drawing.Size(100, 20)
	$rgbLabel.Text = "RGB: $($currentColors.ForeColor)"
	$colorGroup.Controls.Add($rgbLabel)

	$button = New-Object System.Windows.Forms.Button
	$button.Location = New-Object System.Drawing.Point(330, $rowYPos)
	$button.Size = New-Object System.Drawing.Size(110, 25)
	$button.Text = "Choose..."
	$button.Tag = "ForeColor"
	$colorGroup.Controls.Add($button)

	$colorPickerRows["ForeColor"] = @{Panel = $panel; RgbLabel = $rgbLabel; Button = $button; ColorKey = "ForeColor"}
	$rowYPos = $rowYPos + 35

	# Button Background
	$rgb = $currentColors.ButtonBackColor -split ','
	$initialColor = [System.Drawing.Color]::FromArgb([int]$rgb[0], [int]$rgb[1], [int]$rgb[2])

	$labelY = $rowYPos + 5
	$label = New-Object System.Windows.Forms.Label
	$label.Location = New-Object System.Drawing.Point(10, $labelY)
	$label.Size = New-Object System.Drawing.Size(160, 20)
	$label.Text = "Button Background:"
	$colorGroup.Controls.Add($label)

	$panel = New-Object System.Windows.Forms.Panel
	$panel.Location = New-Object System.Drawing.Point(175, $rowYPos)
	$panel.Size = New-Object System.Drawing.Size(40, 25)
	$panel.BorderStyle = "Fixed3D"
	$panel.BackColor = $initialColor
	$panel.Tag = "ColorPickerSwatch"
	$colorGroup.Controls.Add($panel)

	$rgbLabel = New-Object System.Windows.Forms.Label
	$rgbLabel.Location = New-Object System.Drawing.Point(220, $labelY)
	$rgbLabel.Size = New-Object System.Drawing.Size(100, 20)
	$rgbLabel.Text = "RGB: $($currentColors.ButtonBackColor)"
	$colorGroup.Controls.Add($rgbLabel)

	$button = New-Object System.Windows.Forms.Button
	$button.Location = New-Object System.Drawing.Point(330, $rowYPos)
	$button.Size = New-Object System.Drawing.Size(110, 25)
	$button.Text = "Choose..."
	$button.Tag = "ButtonBackColor"
	$colorGroup.Controls.Add($button)

	$colorPickerRows["ButtonBackColor"] = @{Panel = $panel; RgbLabel = $rgbLabel; Button = $button; ColorKey = "ButtonBackColor"}
	$rowYPos = $rowYPos + 35

	# TextBox Background
	$rgb = $currentColors.TextBoxBackColor -split ','
	$initialColor = [System.Drawing.Color]::FromArgb([int]$rgb[0], [int]$rgb[1], [int]$rgb[2])

	$labelY = $rowYPos + 5
	$label = New-Object System.Windows.Forms.Label
	$label.Location = New-Object System.Drawing.Point(10, $labelY)
	$label.Size = New-Object System.Drawing.Size(160, 20)
	$label.Text = "TextBox Background:"
	$colorGroup.Controls.Add($label)

	$panel = New-Object System.Windows.Forms.Panel
	$panel.Location = New-Object System.Drawing.Point(175, $rowYPos)
	$panel.Size = New-Object System.Drawing.Size(40, 25)
	$panel.BorderStyle = "Fixed3D"
	$panel.BackColor = $initialColor
	$panel.Tag = "ColorPickerSwatch"
	$colorGroup.Controls.Add($panel)

	$rgbLabel = New-Object System.Windows.Forms.Label
	$rgbLabel.Location = New-Object System.Drawing.Point(220, $labelY)
	$rgbLabel.Size = New-Object System.Drawing.Size(100, 20)
	$rgbLabel.Text = "RGB: $($currentColors.TextBoxBackColor)"
	$colorGroup.Controls.Add($rgbLabel)

	$button = New-Object System.Windows.Forms.Button
	$button.Location = New-Object System.Drawing.Point(330, $rowYPos)
	$button.Size = New-Object System.Drawing.Size(110, 25)
	$button.Text = "Choose..."
	$button.Tag = "TextBoxBackColor"
	$colorGroup.Controls.Add($button)

	$colorPickerRows["TextBoxBackColor"] = @{Panel = $panel; RgbLabel = $rgbLabel; Button = $button; ColorKey = "TextBoxBackColor"}
	$rowYPos = $rowYPos + 35

	# Gray Label Color
	$rgb = $currentColors.GrayLabelColor -split ','
	$initialColor = [System.Drawing.Color]::FromArgb([int]$rgb[0], [int]$rgb[1], [int]$rgb[2])

	$labelY = $rowYPos + 5
	$label = New-Object System.Windows.Forms.Label
	$label.Location = New-Object System.Drawing.Point(10, $labelY)
	$label.Size = New-Object System.Drawing.Size(160, 20)
	$label.Text = "Gray Label Color:"
	$colorGroup.Controls.Add($label)

	$panel = New-Object System.Windows.Forms.Panel
	$panel.Location = New-Object System.Drawing.Point(175, $rowYPos)
	$panel.Size = New-Object System.Drawing.Size(40, 25)
	$panel.BorderStyle = "Fixed3D"
	$panel.BackColor = $initialColor
	$panel.Tag = "ColorPickerSwatch"
	$colorGroup.Controls.Add($panel)

	$rgbLabel = New-Object System.Windows.Forms.Label
	$rgbLabel.Location = New-Object System.Drawing.Point(220, $labelY)
	$rgbLabel.Size = New-Object System.Drawing.Size(100, 20)
	$rgbLabel.Text = "RGB: $($currentColors.GrayLabelColor)"
	$colorGroup.Controls.Add($rgbLabel)

	$button = New-Object System.Windows.Forms.Button
	$button.Location = New-Object System.Drawing.Point(330, $rowYPos)
	$button.Size = New-Object System.Drawing.Size(110, 25)
	$button.Text = "Choose..."
	$button.Tag = "GrayLabelColor"
	$colorGroup.Controls.Add($button)

	$colorPickerRows["GrayLabelColor"] = @{Panel = $panel; RgbLabel = $rgbLabel; Button = $button; ColorKey = "GrayLabelColor"}
	$rowYPos = $rowYPos + 35

	# Accent Button Color
	$rgb = $currentColors.UpdateButtonColor -split ','
	$initialColor = [System.Drawing.Color]::FromArgb([int]$rgb[0], [int]$rgb[1], [int]$rgb[2])

	$labelY = $rowYPos + 5
	$label = New-Object System.Windows.Forms.Label
	$label.Location = New-Object System.Drawing.Point(10, $labelY)
	$label.Size = New-Object System.Drawing.Size(160, 20)
	$label.Text = "Accent Button Color:"
	$colorGroup.Controls.Add($label)

	$panel = New-Object System.Windows.Forms.Panel
	$panel.Location = New-Object System.Drawing.Point(175, $rowYPos)
	$panel.Size = New-Object System.Drawing.Size(40, 25)
	$panel.BorderStyle = "Fixed3D"
	$panel.BackColor = $initialColor
	$panel.Tag = "ColorPickerSwatch"
	$colorGroup.Controls.Add($panel)

	$rgbLabel = New-Object System.Windows.Forms.Label
	$rgbLabel.Location = New-Object System.Drawing.Point(220, $labelY)
	$rgbLabel.Size = New-Object System.Drawing.Size(100, 20)
	$rgbLabel.Text = "RGB: $($currentColors.UpdateButtonColor)"
	$colorGroup.Controls.Add($rgbLabel)

	$button = New-Object System.Windows.Forms.Button
	$button.Location = New-Object System.Drawing.Point(330, $rowYPos)
	$button.Size = New-Object System.Drawing.Size(110, 25)
	$button.Text = "Choose..."
	$button.Tag = "UpdateButtonColor"
	$colorGroup.Controls.Add($button)

	$colorPickerRows["UpdateButtonColor"] = @{Panel = $panel; RgbLabel = $rgbLabel; Button = $button; ColorKey = "UpdateButtonColor"}

	$yPos = $yPos + 250

	# Preview group box
	$previewGroup = New-Object System.Windows.Forms.GroupBox
	$previewGroup.Location = New-Object System.Drawing.Point(20, $yPos)
	$previewGroup.Size = New-Object System.Drawing.Size(460, 230)
	$previewGroup.Text = "Preview"
	$dialog.Controls.Add($previewGroup)

	# Create preview controls
	$previewPanel = New-Object System.Windows.Forms.Panel
	$previewPanel.Location = New-Object System.Drawing.Point(10, 20)
	$previewPanel.Size = New-Object System.Drawing.Size(440, 200)
	$previewPanel.BorderStyle = "Fixed3D"
	$previewPanel.Tag = "ColorPreviewPanel"  # Mark to skip recursive theme application
	$previewGroup.Controls.Add($previewPanel)

	# Preview label
	$previewLabel = New-Object System.Windows.Forms.Label
	$previewLabel.Location = New-Object System.Drawing.Point(10, 10)
	$previewLabel.Size = New-Object System.Drawing.Size(200, 20)
	$previewLabel.Text = "Sample text label"
	$previewPanel.Controls.Add($previewLabel)

	# Preview gray label
	$previewGrayLabel = New-Object System.Windows.Forms.Label
	$previewGrayLabel.Location = New-Object System.Drawing.Point(10, 40)
	$previewGrayLabel.Size = New-Object System.Drawing.Size(200, 20)
	$previewGrayLabel.Text = "Disabled/gray label"
	$previewPanel.Controls.Add($previewGrayLabel)

	# Preview sample button (in preview panel)
	$previewSampleButton = New-Object System.Windows.Forms.Button
	$previewSampleButton.Location = New-Object System.Drawing.Point(10, 70)
	$previewSampleButton.Size = New-Object System.Drawing.Size(120, 30)
	$previewSampleButton.Text = "Sample Button"
	$previewSampleButton.FlatStyle = "Standard"  # Use standard 3D style that respects BackColor
	$previewPanel.Controls.Add($previewSampleButton)

	# Preview accent button (in preview panel)
	$previewAccentButton = New-Object System.Windows.Forms.Button
	$previewAccentButton.Location = New-Object System.Drawing.Point(140, 70)
	$previewAccentButton.Size = New-Object System.Drawing.Size(140, 30)
	$previewAccentButton.Text = "Accent Button"
	$previewAccentButton.FlatStyle = "Standard"  # Use standard 3D style that respects BackColor
	$previewPanel.Controls.Add($previewAccentButton)

	# Preview textbox
	$previewTextBox = New-Object System.Windows.Forms.TextBox
	$previewTextBox.Location = New-Object System.Drawing.Point(10, 110)
	$previewTextBox.Size = New-Object System.Drawing.Size(270, 25)
	$previewTextBox.Text = "Sample text input"
	$previewPanel.Controls.Add($previewTextBox)

	# Preview combobox
	$previewCombo = New-Object System.Windows.Forms.ComboBox
	$previewCombo.Location = New-Object System.Drawing.Point(10, 145)
	$previewCombo.Size = New-Object System.Drawing.Size(270, 25)
	$previewCombo.Items.AddRange(@("Option 1", "Option 2", "Option 3"))
	$previewCombo.SelectedIndex = 0
	$previewPanel.Controls.Add($previewCombo)

	# Function to update preview with current colors
	function global:Update-ColorPreview {
		param($ColorValues)

		# Parse colors
		$backRgb = $ColorValues.BackColor -split ','
		$foreRgb = $ColorValues.ForeColor -split ','
		$buttonBackRgb = $ColorValues.ButtonBackColor -split ','
		$textBoxBackRgb = $ColorValues.TextBoxBackColor -split ','
		$grayRgb = $ColorValues.GrayLabelColor -split ','
		$accentRgb = $ColorValues.UpdateButtonColor -split ','

		$backColor = [System.Drawing.Color]::FromArgb([int]$backRgb[0], [int]$backRgb[1], [int]$backRgb[2])
		$foreColor = [System.Drawing.Color]::FromArgb([int]$foreRgb[0], [int]$foreRgb[1], [int]$foreRgb[2])
		$buttonBackColor = [System.Drawing.Color]::FromArgb([int]$buttonBackRgb[0], [int]$buttonBackRgb[1], [int]$buttonBackRgb[2])
		$textBoxBackColor = [System.Drawing.Color]::FromArgb([int]$textBoxBackRgb[0], [int]$textBoxBackRgb[1], [int]$textBoxBackRgb[2])
		$grayColor = [System.Drawing.Color]::FromArgb([int]$grayRgb[0], [int]$grayRgb[1], [int]$grayRgb[2])
		$accentColor = [System.Drawing.Color]::FromArgb([int]$accentRgb[0], [int]$accentRgb[1], [int]$accentRgb[2])

		# Apply to preview controls
		$previewPanel.BackColor = $backColor
		$previewLabel.ForeColor = $foreColor
		$previewGrayLabel.ForeColor = $grayColor

		# Set button colors (FlatStyle = Standard respects BackColor)
		$previewSampleButton.BackColor = $buttonBackColor
		$previewSampleButton.ForeColor = $foreColor
		$previewSampleButton.Refresh()

		$previewAccentButton.BackColor = $accentColor
		$previewAccentButton.ForeColor = $foreColor
		$previewAccentButton.Refresh()

		$previewTextBox.BackColor = $textBoxBackColor
		$previewTextBox.ForeColor = $foreColor

		$previewCombo.BackColor = $textBoxBackColor
		$previewCombo.ForeColor = $foreColor
	}

	# Initial preview update
	Update-ColorPreview -ColorValues $currentColors

	# Handle color picker button clicks
	foreach ($key in $colorPickerRows.Keys) {
		$row = $colorPickerRows[$key]
		$row.Button.Add_Click({
			param($eventSender, $e)

			$colorKey = $eventSender.Tag
			$currentRow = $colorPickerRows[$colorKey]

			# Open ColorDialog
			$colorDialog = New-Object System.Windows.Forms.ColorDialog
			$colorDialog.AllowFullOpen = $true
			$colorDialog.FullOpen = $true
			$colorDialog.Color = $currentRow.Panel.BackColor

			if ($colorDialog.ShowDialog() -eq 'OK') {
				$selectedColor = $colorDialog.Color
				$rgbString = "$($selectedColor.R),$($selectedColor.G),$($selectedColor.B)"

				# Update panel and label
				$currentRow.Panel.BackColor = $selectedColor
				$currentRow.RgbLabel.Text = "RGB: $rgbString"

				# Update preview
				$tempColors = @{
					BackColor = $colorPickerRows["BackColor"].RgbLabel.Text -replace 'RGB: ', ''
					ForeColor = $colorPickerRows["ForeColor"].RgbLabel.Text -replace 'RGB: ', ''
					ButtonBackColor = $colorPickerRows["ButtonBackColor"].RgbLabel.Text -replace 'RGB: ', ''
					TextBoxBackColor = $colorPickerRows["TextBoxBackColor"].RgbLabel.Text -replace 'RGB: ', ''
					GrayLabelColor = $colorPickerRows["GrayLabelColor"].RgbLabel.Text -replace 'RGB: ', ''
					UpdateButtonColor = $colorPickerRows["UpdateButtonColor"].RgbLabel.Text -replace 'RGB: ', ''
				}
				Update-ColorPreview -ColorValues $tempColors
			}
		}.GetNewClosure())
	}

	# Handle preset selection
	$presetCombo.Add_SelectedIndexChanged({
		$selectedPreset = $presetCombo.SelectedItem

		if ($selectedPreset -ne "Current Colors" -and $presets.ContainsKey($selectedPreset)) {
			$presetColors = $presets[$selectedPreset]

			# Update all color pickers
			foreach ($key in $presetColors.Keys) {
				$rgb = $presetColors[$key] -split ','
				$color = [System.Drawing.Color]::FromArgb([int]$rgb[0], [int]$rgb[1], [int]$rgb[2])

				$colorPickerRows[$key].Panel.BackColor = $color
				$colorPickerRows[$key].RgbLabel.Text = "RGB: $($presetColors[$key])"
			}

			# Update preview
			Update-ColorPreview -ColorValues $presetColors
		}
	}.GetNewClosure())

	$yPos = $yPos + 240

	# Bottom buttons
	$buttonY = $yPos + 10
	$buttonWidth = 90
	$buttonSpacing = 10

	# Save Theme button event handler
	$btnSaveTheme.Add_Click({
		# Show themed input dialog for theme name
		$themeName = Show-ThemedInputDialog -Title "Save Theme" -Prompt "Enter a name for this theme:" -DefaultValue "My Custom Theme" -ParentIcon $dialog.Icon

		if ([string]::IsNullOrWhiteSpace($themeName)) {
			return  # User canceled or empty name
		}

		# Check if file already exists
		$safeThemeName = $themeName -replace '[\\/:*?"<>|]', '_'
		$potentialPath = Join-Path $themesDir "$safeThemeName.json"

		if (Test-Path $potentialPath) {
			# Show themed overwrite confirmation dialog
			$result = Show-ThemedMessageDialog -Title "File Exists" -Message "A theme file with this name already exists:`n$potentialPath`n`nDo you want to overwrite it?" -Buttons "OKCancel" -Icon "Warning" -ParentIcon $dialog.Icon
			if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
				return  # User chose not to overwrite
			}
		}

		# Collect current colors from UI
		$colorsToSave = @{
			BackColor = $colorPickerRows["BackColor"].RgbLabel.Text -replace 'RGB: ', ''
			ForeColor = $colorPickerRows["ForeColor"].RgbLabel.Text -replace 'RGB: ', ''
			ButtonBackColor = $colorPickerRows["ButtonBackColor"].RgbLabel.Text -replace 'RGB: ', ''
			TextBoxBackColor = $colorPickerRows["TextBoxBackColor"].RgbLabel.Text -replace 'RGB: ', ''
			GrayLabelColor = $colorPickerRows["GrayLabelColor"].RgbLabel.Text -replace 'RGB: ', ''
			UpdateButtonColor = $colorPickerRows["UpdateButtonColor"].RgbLabel.Text -replace 'RGB: ', ''
		}

		# Export theme
		$savedPath = Export-SandboxStartTheme -Colors $colorsToSave -ThemeName $themeName

		if ($savedPath) {
			# Show themed success dialog
			Show-ThemedMessageDialog -Title "Theme Saved" -Message "Theme saved successfully to:`n$savedPath" -Buttons "OK" -Icon "Information" -ParentIcon $dialog.Icon | Out-Null
		} else {
			# Show themed error dialog
			Show-ThemedMessageDialog -Title "Save Error" -Message "Failed to save theme. Check error messages for details." -Buttons "OK" -Icon "Error" -ParentIcon $dialog.Icon | Out-Null
		}
	}.GetNewClosure())

	# Load Theme button event handler
	$btnLoadTheme.Add_Click({
		# Create OpenFileDialog
		$openDialog = New-Object System.Windows.Forms.OpenFileDialog
		$openDialog.Title = "Load Theme"
		$openDialog.Filter = "JSON Theme Files (*.json)|*.json|All Files (*.*)|*.*"

		# Set initial directory to themes folder (create if doesn't exist)
		if (-not (Test-Path $themesDir)) {
			New-Item -ItemType Directory -Path $themesDir -Force | Out-Null
		}
		$openDialog.InitialDirectory = $themesDir

		if ($openDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
			# Import theme
			$importedTheme = Import-SandboxStartTheme -FilePath $openDialog.FileName

			if ($importedTheme -and $importedTheme.Colors) {
				# Update all color pickers with loaded values
				foreach ($key in $importedTheme.Colors.Keys) {
					if ($colorPickerRows.ContainsKey($key)) {
						$rgb = $importedTheme.Colors[$key] -split ','
						$color = [System.Drawing.Color]::FromArgb([int]$rgb[0], [int]$rgb[1], [int]$rgb[2])

						$colorPickerRows[$key].Panel.BackColor = $color
						$colorPickerRows[$key].RgbLabel.Text = "RGB: $($importedTheme.Colors[$key])"
					}
				}

				# Update preview
				Update-ColorPreview -ColorValues $importedTheme.Colors

				# Reset preset combo to "Current Colors"
				$presetCombo.SelectedIndex = 0
			} else {
				[System.Windows.Forms.MessageBox]::Show(
					"Failed to load theme. The file may be invalid or corrupted.",
					"Load Error",
					[System.Windows.Forms.MessageBoxButtons]::OK,
					[System.Windows.Forms.MessageBoxIcon]::Error
				)
			}
		}
	}.GetNewClosure())

	# Reset button
	$resetButton = New-Object System.Windows.Forms.Button
	$resetButton.Location = New-Object System.Drawing.Point(20, $buttonY)
	$resetButton.Size = New-Object System.Drawing.Size($buttonWidth, 30)
	$resetButton.Text = "Reset"
	$resetButton.Add_Click({
		# Find "Visual Studio Dark (Default)" in the combo
		$defaultIndex = -1
		for ($i = 0; $i -lt $presetCombo.Items.Count; $i++) {
			if ($presetCombo.Items[$i] -eq "Visual Studio Dark (Default)") {
				$defaultIndex = $i
				break
			}
		}

		if ($defaultIndex -ge 0) {
			# If already selected, manually update colors
			if ($presetCombo.SelectedIndex -eq $defaultIndex) {
				$defaultColors = $presets["Visual Studio Dark (Default)"]

				# Update all color pickers
				foreach ($key in $defaultColors.Keys) {
					$rgb = $defaultColors[$key] -split ','
					$color = [System.Drawing.Color]::FromArgb([int]$rgb[0], [int]$rgb[1], [int]$rgb[2])

					$colorPickerRows[$key].Panel.BackColor = $color
					$colorPickerRows[$key].RgbLabel.Text = "RGB: $($defaultColors[$key])"
				}

				# Update preview
				Update-ColorPreview -ColorValues $defaultColors
			} else {
				# Just change selection, event handler will update
				$presetCombo.SelectedIndex = $defaultIndex
			}
		}
	}.GetNewClosure())
	$dialog.Controls.Add($resetButton)

	# Preview button (temporarily apply without saving)
	$previewButton = New-Object System.Windows.Forms.Button
	$previewX = 20 + $buttonWidth + $buttonSpacing
	$previewButton.Location = New-Object System.Drawing.Point($previewX, $buttonY)
	$previewButton.Size = New-Object System.Drawing.Size($buttonWidth, 30)
	$previewButton.Text = "Preview"
	$previewButton.Add_Click({
		# Collect current colors from UI
		$newColors = @{
			BackColor = $colorPickerRows["BackColor"].RgbLabel.Text -replace 'RGB: ', ''
			ForeColor = $colorPickerRows["ForeColor"].RgbLabel.Text -replace 'RGB: ', ''
			ButtonBackColor = $colorPickerRows["ButtonBackColor"].RgbLabel.Text -replace 'RGB: ', ''
			TextBoxBackColor = $colorPickerRows["TextBoxBackColor"].RgbLabel.Text -replace 'RGB: ', ''
			GrayLabelColor = $colorPickerRows["GrayLabelColor"].RgbLabel.Text -replace 'RGB: ', ''
			UpdateButtonColor = $colorPickerRows["UpdateButtonColor"].RgbLabel.Text -replace 'RGB: ', ''
		}

		# Apply to parent form using custom theme (but don't save to registry yet)
		$updateRgb = $newColors.UpdateButtonColor -split ','
		$updateColor = [System.Drawing.Color]::FromArgb([int]$updateRgb[0], [int]$updateRgb[1], [int]$updateRgb[2])
		Set-CustomTheme -Control $ParentForm -CustomColors $newColors -UpdateButtonBackColor $updateColor

		# Update title bar based on background brightness
		$bgRgb = $newColors.BackColor -split ','
		$bgColor = [System.Drawing.Color]::FromArgb([int]$bgRgb[0], [int]$bgRgb[1], [int]$bgRgb[2])
		$isDark = Test-ColorIsDark -Color $bgColor
		Set-DarkTitleBar -Form $ParentForm -UseDarkMode $isDark
	}.GetNewClosure())
	$dialog.Controls.Add($previewButton)

	# OK button (aligned with preview box, second from right)
	$okButton = New-Object System.Windows.Forms.Button
	$okX = 20 + 460 - (2 * $buttonWidth) - $buttonSpacing
	$okButton.Location = New-Object System.Drawing.Point($okX, $buttonY)
	$okButton.Size = New-Object System.Drawing.Size($buttonWidth, 30)
	$okButton.Text = "OK"
	$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$okButton.Add_Click({
		# Same as Apply, but also closes dialog
		$newColors = @{
			BackColor = $colorPickerRows["BackColor"].RgbLabel.Text -replace 'RGB: ', ''
			ForeColor = $colorPickerRows["ForeColor"].RgbLabel.Text -replace 'RGB: ', ''
			ButtonBackColor = $colorPickerRows["ButtonBackColor"].RgbLabel.Text -replace 'RGB: ', ''
			TextBoxBackColor = $colorPickerRows["TextBoxBackColor"].RgbLabel.Text -replace 'RGB: ', ''
			GrayLabelColor = $colorPickerRows["GrayLabelColor"].RgbLabel.Text -replace 'RGB: ', ''
			UpdateButtonColor = $colorPickerRows["UpdateButtonColor"].RgbLabel.Text -replace 'RGB: ', ''
		}

		Set-SandboxStartCustomColors -Colors $newColors
		Set-SandboxStartThemePreference -ThemeMode "Custom"

		$updateRgb = $newColors.UpdateButtonColor -split ','
		$updateColor = [System.Drawing.Color]::FromArgb([int]$updateRgb[0], [int]$updateRgb[1], [int]$updateRgb[2])
		Set-CustomTheme -Control $ParentForm -CustomColors $newColors -UpdateButtonBackColor $updateColor

		# Update title bar based on background brightness
		$bgRgb = $newColors.BackColor -split ','
		$bgColor = [System.Drawing.Color]::FromArgb([int]$bgRgb[0], [int]$bgRgb[1], [int]$bgRgb[2])
		$isDark = Test-ColorIsDark -Color $bgColor
		Set-DarkTitleBar -Form $ParentForm -UseDarkMode $isDark

		# Refresh context menu to show Custom checkmark
		$ParentForm.ContextMenuStrip = Show-ThemeContextMenu -Form $ParentForm -UpdateButtonColor $UpdateButtonColor -WorkingDir $localWorkingDir -TestRegKey $localTestRegKey -TestContextMenu $localTestContextMenu -UpdateContextMenu $localUpdateContextMenu -AppIcon $localAppIcon
	}.GetNewClosure())
	$dialog.Controls.Add($okButton)

	# Cancel button (restore original theme) (aligned with preview box right edge)
	$cancelButton = New-Object System.Windows.Forms.Button
	$cancelX = 20 + 460 - $buttonWidth
	$cancelButton.Location = New-Object System.Drawing.Point($cancelX, $buttonY)
	$cancelButton.Size = New-Object System.Drawing.Size($buttonWidth, 30)
	$cancelButton.Text = "Cancel"
	$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$cancelButton.Add_Click({
		# Restore original theme
		if ($originalTheme -eq "Custom") {
			$updateRgb = $originalColors.UpdateButtonColor -split ','
			$updateColor = [System.Drawing.Color]::FromArgb([int]$updateRgb[0], [int]$updateRgb[1], [int]$updateRgb[2])
			Set-CustomTheme -Control $ParentForm -CustomColors $originalColors -UpdateButtonBackColor $updateColor

			$bgRgb = $originalColors.BackColor -split ','
			$bgColor = [System.Drawing.Color]::FromArgb([int]$bgRgb[0], [int]$bgRgb[1], [int]$bgRgb[2])
			$isDark = Test-ColorIsDark -Color $bgColor
			Set-DarkTitleBar -Form $ParentForm -UseDarkMode $isDark
		} else {
			Set-ThemeToForm -Form $ParentForm -UpdateButtonColor $UpdateButtonColor
		}

		# Restore context menu
		$ParentForm.ContextMenuStrip = Show-ThemeContextMenu -Form $ParentForm -UpdateButtonColor $UpdateButtonColor -WorkingDir $localWorkingDir -TestRegKey $localTestRegKey -TestContextMenu $localTestContextMenu -UpdateContextMenu $localUpdateContextMenu -AppIcon $localAppIcon
	}.GetNewClosure())
	$dialog.Controls.Add($cancelButton)

	# Set accept/cancel buttons
	$dialog.AcceptButton = $okButton
	$dialog.CancelButton = $cancelButton

	# Apply current theme to the dialog
	$currentTheme = Get-SandboxStartThemePreference
	if ($currentTheme -eq "Custom") {
		$customColors = Get-SandboxStartCustomColors
		$updateRgb = $customColors.UpdateButtonColor -split ','
		$themeUpdateColor = [System.Drawing.Color]::FromArgb([int]$updateRgb[0], [int]$updateRgb[1], [int]$updateRgb[2])
		Set-CustomTheme -Control $dialog -CustomColors $customColors -UpdateButtonBackColor $themeUpdateColor

		# Set title bar color based on background brightness
		$bgRgb = $customColors.BackColor -split ','
		$bgColor = [System.Drawing.Color]::FromArgb([int]$bgRgb[0], [int]$bgRgb[1], [int]$bgRgb[2])
		$isDark = Test-ColorIsDark -Color $bgColor
		Set-DarkTitleBar -Form $dialog -UseDarkMode $isDark
	} else {
		Set-ThemeToForm -Form $dialog -UpdateButtonColor $UpdateButtonColor
	}

	# Show dialog
	$dialog.ShowDialog() | Out-Null
}

# Define the dialog function here since it's needed before the main functions section
function Show-SandboxTestDialog {
	<#
	.SYNOPSIS
	Shows a GUI dialog for configuring Windows Sandbox test parameters

	.DESCRIPTION
	Creates a Windows Forms dialog to collect all parameters needed for SandboxTest function.
	Can pre-fill paths from parent script via $script:InitialFolderPath and $script:InitialFilePath.
	#>

	# Embedded icon data (Base64-encoded from Source\assets\icon.ico)
	# Generated: 2025-12-20
	# Original size: 15KB (.ico) ??? 20KB (base64)
	$iconBase64 = @"
AAABAAMAMDAAAAEAIACoJQAANgAAACAgAAABACAAqBAAAN4lAAAQEAAAAQAgAGgEAACGNgAAKAAAADAAAABgAAAAAQAgAAAAAACAJQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA/78/BP+/PwQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP/MMw/+yzNP/sgun/7EKp3+wCVK/rAnDQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD+2jYO/tE6T/7RN7P+zTPy/8gt///CKP//viLw/rserv63G0v+sBMNAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA/to2Dv7ROk/+0jiz/tM58f/ROP//zjP//8gt///CJ///vSH//7kd//+3G/D+uBqu/rYbSv6wEw0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP7aNg7+0TpP/tI6s/7TOvH/0jn//9I4///RN///zTL//8cs///BJv//vCD//7gc//+3G//+thr+/7UZ8P60GK/+thhK/rATDQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD+2jYO/tE6T/7UOrP+0zrx/9M6///SOf//0jj//9I4///RN///zTL//8cs///BJv//vCD//7gb//+2Gv//thn//7UY//+0F///sxfw/rMVr/6zFEr+sBMNAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA/to2Dv7ROk/+1Dqz/tQ68f/TOv//0zn//9I5///SOP//0jj//9I4///RN///zTP//8gt///CJ///vSH//7gc//+2Gv//tRn//7QY//+0F///sxb//rIV/v+xFPD+sRSu/rMRSv6wEw0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP/MMw/+1TpP/tQ6s/7TO/L+0zr+/9M5///TOf//0zn//9I5///SOf//0jj//9I5///SOP//zjP//8gt///DKP//vSL//7kd//+3G///thn//7UY//+zF///shX//7EU//+wE///sBL//68S8P6uEa/+rxFK/rATDQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD+2jYO/tU6T/7UO7P+0zvy/9M7///TOv//0zr//9M5///TOf//0zn//9I5///TOf//0zn//9M5///SOP//zjT//8ku///EKf//vyP//7oe//+4HP//txr//7UZ//+0F///shX//7EU//+wE///rxL//64R//6tEP7/rQ/w/qwOrv6sDUr+sBMNAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA/to2Dv7VPU/+1Duz/tM78v/UO///0zr//9M6///TOv//0zr//9M6///TOf//0zr//9M6///TOv//0zr//9Q6///TOv//0DX//8sw///GK///wSX//7wg//+6Hv//uBz//7Ya//+1GP//sxb//7IU//+wE///rxL//64Q//+tD///rA7//6sO//+rDfD+qg2v/qwNSv6cEw0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP/MRA/+1T1P/tU7s/7UPPL/1Dz//9Q7///TOv//0zr//9M6///TOv//0zr//9M6///UOv//0zr//9Q6///UO///1Dv//9U7///UO///0Tf//8wy///ILf//wyj//74j//+8If//uh7//7gc//+2Gv//tBj//7MW//+xE///rxL//64R//+tD///rA7//6sN//+qDP//qQv//6gK8P6nCq/+qApK/rATDQAAAAAAAAAAAAAAAAAAAAD+2kgO/tU9T/7VPbP+1Dzy/9Q8///UO///1Dv//9Q7///UO///0zr//9M6///TOv//0zv//9Q6///UO///1Dv//9Q7//7VPP/+1Tz//tU8///VPP//0jn//840//7KL//+xSr//8Em//6/JP//vCH//7of//+4HP//thr//7QX//+yFf//sBP//64R//+tD///rA7//6sN//+qC///qQr//6gJ//+nCf//pwjw/qcIr/6mCkv+sBMNAAAAAP/XP0D+1z6z/tU+8v/VPf//1Tz//9Q7///UO///1Dv//9Q7///UOv//1Dv//9Q7//7UO//+1Dv//tQ7///UPP/+1Tz//9U8//7WPf/+1j3//tY+///WPv//0zr//c41//7KMf/+xy3//sMp///BJv//vyT//70i//+6H//+uBz//rYa//+0F///shX//7AT//+uEf//rQ///6sN//+qDP//qAr//6cJ//+nCP//pgj//6cI//+oCfD+qQuu/qwMPv7XQNv/1j7//9Y+///VPf//1Dz//9Q7///UO///1Dv//9Q7///UO///1Dv//9Q7///UPP/+1Tz//9U8//7WPf//1j3//9Y+//7WPv/+1z///tc///vSPP/uuyv/3Z8Y/92dF//vsyL//MIp///EKf//wif//sAl//69Iv//uyD//rkd//+2Gv//tBj//rIV//+wEv//rhH//6wO//+qDP//qQr//6gJ//+nCP//pgf//6YI//+nCf/+qQv+/qsO2v/YQP//1z///9Y9///VPf//1Tz//9Q7///UO///1Dz//9U8///VPP//1Tz//9U8///VPf//1j3//tY+//7XPv//1z7//9c////YP//80z3/774v/9qcGP/Ohgn/zIUM/8uNIP/PixX/25gU/++xIP/7wCj//sIn//7AJf//viP//7wg//65Hf//txv//rUY//+yFf//sBP//64Q//+sDv//qgz//6gK//+nCP//pgf//6YH//+nCf//qQv//qwO/v/YQP//1z///9Y+///VPf//1T3//9U8///VPP//1Tz//9U9///VPf//1j3//9Y+///WPv//1z7//9c////YP///2ED//NQ+/+/JSP/Jwnj/vqtn/86NHP/QiA//vqhj/3+9wP+KuK3/taZo/8+IDf/blxL/77Af//y/Jv//wCb//74j//68If//uh7//7gb//+1Gf//sxb//7AT//+uEf//rA7//6oM//+oCv//pwj//6cI//+oCf//qQv//6wO///YQP//1z///9Y+///VPf//1T3//9U8///VPf//1j3//9Y+///XPv//1z7//9c///7XP///2ED//9hA//zUPf/wvzD/3Ko2/5rIuP9dxN//WbzY/4i2rv+btZ7/dMbV/02+3/9Lutv/c7e+/8uSKv/Ogwb/z4YI/9yWEf/wrx7//L4l//6/JP//vSH//7sf//+4HP//thr//7MX//+xE///rhH//6wO//+qDP//qQr//6gJ//+pCv//qgz//6wP///YQP//1z///9Y+///WPv//1j7//9Y9///WPv//1z7//9c+///XP///2D///9hA//7YQP/81D7/8cEx/9+hHf/UjQ//wapi/2fT6f9Tz+z/Sr/h/0m32/9Lv+D/V9z1/1jd9v9Rzur/TbXT/5q0nP/FoE3/w5xG/8mWM//RiQ3/3JgT//CwHv/8vCP//r0i//+7IP//uR3//7ca//+0F///sRT//68R//+sD///qw3//6oM//+qDP//qw3//60Q///YQP//2D///9c////XP///1z7//9c////XP///2D///9hA///YQP//2UH//dU+//HDM//hpSH/15IT/9WOEf/Wkx3/r8Wb/1/n/P9b5v3/V973/1XW8f9Y4Pf/WOj+/1jo/v9V4vv/RL3e/0m11f9ct8//WrjR/3m7wP/IoUv/0YkM/9ONDv/emxX/8bAe//y7Iv//vCD//7oe//+3G///tRj//7IV//+wEv//rhD//60P//+s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AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJxOEw2hVgpKnlIJnptQCJ+ZTgZOkUgADgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAvz8ABL8/AAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP///////wAA////////AAD///////8AAP///////wAA///+f///AAD///gf//8AAP//4Af//wAA//+AAf//AAD//gAAf/8AAP/4AAAf/wAA/+AAAAf/AAD/gAAAAf8AAP4AAAAAfwAA+AAAAAAfAADgAAAAAAcAAIAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACAAAAAAAEAAOAAAAAABwAA+AAAAAAfAAD+AAAAAH8AAP+AAAAB/wAA/+AAAAf/AAD/+AAAH/8AAP/+AAB//wAA//+AAf//AAD//+AH//8AAP//+B///wAA///+f///AAD///////8AAP///////wAA////////AAD///////8AACgAAAAgAAAAQAAAAAEAIAAAAAAAgBAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//8AAf7RNhz/yS9g/sMoX/68JRv//wABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//8AAf7RNhz+0zlv/tA2z/7ILvr+wSX5/7sfzP64HGz+sxwb//8AAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//8AAf7aPxz+0zlv/tM4zv7SOfr/0Db//8ku///AJf//uR3//rca+f+2Gsz+tRds/rMSG///AAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//8AAf7aPxz+0zlv/tI5z/7TOfr/0jn//9I4///QNv//yC7//8Ak//+5Hf//thr//7UZ//60F/n/shbM/rEVbP6zEhv//wABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//8AAf7aPxz+1Ttv/tM5z/7TOvr/0zn//9I5///SOP//0jj//9E3///JL///wSb//7oe//+3Gv//tRj//7MW//+yFf/+sRP5/7ASzP6uEGz+qRIb//8AAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//8AAf7aPxz+1Ttv/tM7z/7TO/r/0zr//9M5///TOf//0zn//9M5///TOf//0Tj//8sx///DKP//vCD//7gc//+2Gf//tBf//7IU//+wE///rxH//q4P+f+sD8z+rA5s/qkJG///AAEAAAAAAAAAAAAAAAAAAAAA//8AAf7aPxz+1Ttv/tM8z/7UO/r/1Dr//9M6///TOv//0zr//9M6///TOv//1Dr//9Q7///TOf//zTP//8Yr//+/I///ux///7gc//+1Gf//sxb//7AT//+uEf//rQ///6sN//6rDPn/qgvM/qcJbP6pCRv//wABAAAAAP7ZQhv+1T5v/tU8z/7UPPr/1Dv//9Q7///UOv//1Dv//9Q7///UO///1Dv//9Q8///VPP//1T3//9U8///PNv//yS///sIo//+/JP//vCD//7gc//+1Gf//shX//68S//+tD///qw3//6kL//+oCf/+pwj5/6cIzP6nCWz+rQoZ/tc+w/7WPvr/1T3//9Q7///UO///1Dv//9Q7///UO///1Dz//9U8//7VPf/+1j3//tY+//7XPv/5zTf/67Ml/+uvIf/5vif//sIo///AJP//vCH//7kd//+1Gf//shX//68S//+sD///qgz//6gJ//+mCP//pgj//qgK+f6rDcH+1z/+/9Y+///VPf//1Dz//9Q8///VPP//1Tz//9U9///WPf//1j7//9c+//7XP//50kL/5r0//9aUFP/LjBz/uptJ/82cMf/oqR3/+bwl//7AJf//vSL//7ke//+2Gf//sxb//68S//+sDv//qQv//6cI//+mCP//qAn//qsN/v/XP///1j7//9U9///VPP//1T3//9Y9///WPv//1z7//9c///7XP//5zzr/6Lw8/6TFoP94u7z/oKh8/5K4ov9bwNn/erSw/8qPI//Wjw3/6aca//m6Iv/+vSL//7of//+3G///sxb//7AS//+sDv//qQv//6gJ//+pCv//qw7//9hA///XP///1j7//9Y+///WPv//1z///9g////YQP/60Dv/7Lgr/9uZGP/Ap1X/Zdjr/1DM6v9Pw+L/Vtnz/1bd9v9Pv9v/i7Cc/6amcv+8nUv/2JMT/+qoGv/6uSH//rsg//+4HP//tBf//7AT//+tD///qw3//6sN//+tD///2UH//9hA///XP///1z///9hA//7YQP/3zjv/7bov/9+hHv/ZlBX/xaVQ/43G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
"@

	try {
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

		# Initialize sandbox configuration file
		Initialize-SandboxConfig -WorkingDir $Script:WorkingDir

		# Initialize package list migration (one-time tracking)
		Initialize-PackageListMigration -WorkingDir $Script:WorkingDir

		# Ensure AutoInstall.txt exists
		$autoInstallPath = Join-Path (Join-Path $Script:WorkingDir "wsb") "AutoInstall.txt"
		if (-not (Test-Path $autoInstallPath)) {
			$autoInstallContent = @"
# AutoInstall Package List (Local Only)
# This list is installed FIRST, before any selected package lists
# Add WinGet package IDs (one per line) to install automatically
# Example: Notepad++.Notepad++
#
# Usage:
# - Always runs first (even if not selected)
# - Can be manually selected to ONLY install these packages
# - Cannot be deleted or synced from GitHub

"@
			Set-Content -Path $autoInstallPath -Value $autoInstallContent -Encoding UTF8
			Write-Verbose "Created AutoInstall.txt"
		}

		# Download/update default scripts from GitHub
		Write-Host "Checking default scripts...`t" -NoNewline -ForegroundColor Cyan
		$initialStatus = "Checking default scripts from GitHub"
		Sync-GitHubScriptsSelective -LocalFolder $wsbDir -UseCache
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
		$form.Size = New-Object System.Drawing.Size(465, 770)
		$form.StartPosition = "CenterScreen"
		$form.FormBorderStyle = "FixedDialog"
		$form.MaximizeBox = $false
		$form.MinimizeBox = $true

		# Set custom icon if available
		if ($appIcon) {
			$form.Icon = $appIcon
			$form.ShowIcon = $true
		}
		else {
			$form.ShowIcon = $false
		}

		# Define adaptive green color for Update button based on theme
		# Note: Will be determined after theme is applied
		$themeMode = Get-SandboxStartThemePreference
		$tempDark = if ($themeMode -eq "Auto") {
			Get-WindowsThemeSetting
		} elseif ($themeMode -eq "Dark" -or $themeMode -eq "Custom") {
			$true
		} else {
			$false
		}

		$updateButtonGreen = if ($tempDark) {
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
		$controlWidth = 409

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

		# Read-Only notification label (centered above mapped folder controls)
		$lblReadOnlyInfo = New-Object System.Windows.Forms.Label
		$lblReadOnlyInfo.Location = New-Object System.Drawing.Point(($leftMargin + ($controlWidth / 2) - 57), $y)
		$lblReadOnlyInfo.Size = New-Object System.Drawing.Size(100, 15)
		$lblReadOnlyInfo.Text = "R/O by default!"
		$lblReadOnlyInfo.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
		$lblReadOnlyInfo.ForeColor = [System.Drawing.SystemColors]::HotTrack
		$lblReadOnlyInfo.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
		$form.Controls.Add($lblReadOnlyInfo)

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
		$txtMapFolder.Enabled = $false  # Disable direct editing - users must use browse buttons

		$y += $labelHeight + $controlHeight + 5

		# Folder browse button
		$btnBrowse = New-Object System.Windows.Forms.Button
		$btnBrowse.Location = New-Object System.Drawing.Point($leftMargin, $y)
		$btnBrowse.Size = New-Object System.Drawing.Size(($controlWidth * 0.44), $controlHeight)
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

				# Use shared function to update form (folder)
				
Update-FormFromSelection -SelectedPath $selectedDir -txtMapFolder $txtMapFolder -txtSandboxFolderName $txtSandboxFolderName -txtScript $txtScript -lblStatus $lblStatus -btnSaveScript $btnSaveScript -chkNetworking $chkNetworking -chkSkipWinGet $chkSkipWinGet -cmbInstallPackages $cmbInstallPackages -wsbDir $wsbDir
			}
		})
		$form.Controls.Add($btnBrowse)

		# Read-Only checkbox in middle (small, with label below)
		$chkMapFolderReadOnly = New-Object System.Windows.Forms.CheckBox
		$chkMapFolderReadOnly.Location = New-Object System.Drawing.Point(($leftMargin + $controlWidth * 0.44 + 10), ($y + 2))
		$chkMapFolderReadOnly.Size = New-Object System.Drawing.Size(15, 15)
		$chkMapFolderReadOnly.Text = ""
		$chkMapFolderReadOnly.Checked = $true  # Default to read-only
		$tooltipReadOnly = New-Object System.Windows.Forms.ToolTip
		$tooltipReadOnly.SetToolTip($chkMapFolderReadOnly, "Map the folder as read-only in the sandbox. Prevents any modifications to source files.")
		$form.Controls.Add($chkMapFolderReadOnly)

		# R/O label below checkbox
		$lblReadOnly = New-Object System.Windows.Forms.Label
		$lblReadOnly.Location = New-Object System.Drawing.Point(($leftMargin + $controlWidth * 0.44 + 4), ($y + 17))
		$lblReadOnly.Size = New-Object System.Drawing.Size(25, 15)
		$lblReadOnly.Text = "R/O"
		$lblReadOnly.Font = New-Object System.Drawing.Font("Segoe UI", 7)
		$lblReadOnly.TextAlign = [System.Drawing.ContentAlignment]::TopCenter
		$form.Controls.Add($lblReadOnly)

		# File browse button (aligned to right edge of control width)
		$btnBrowseFile = New-Object System.Windows.Forms.Button
		$btnBrowseFile.Location = New-Object System.Drawing.Point(($leftMargin + $controlWidth * 0.44 + 35), $y)
		$btnBrowseFile.Size = New-Object System.Drawing.Size(($controlWidth - ($controlWidth * 0.44) - 35), $controlHeight)
		$btnBrowseFile.Text = "File..."
		$btnBrowseFile.Add_Click({
			$fileDialog = New-Object System.Windows.Forms.OpenFileDialog
			$fileDialog.Title = "Select file to run in Windows Sandbox"

			# Build dynamic filter from INI extensions
			$baseExtensions = @("exe", "msi", "msix", "cmd", "bat", "ps1", "appx", "appxbundle", "intunewin")
			$defaultScriptExtensions = @("ahk", "au3", "py", "js")

			# Load custom extensions from INI
			$extensionMappings = Get-SandboxConfig -Section 'Extensions' -WorkingDir $wsbDir
			$customExtensions = @()

			foreach ($ext in ($extensionMappings.Keys | Sort-Object)) {
				# Exclude base executables and default script extensions
				if ($ext -notin $baseExtensions -and $ext -notin $defaultScriptExtensions) {
					$customExtensions += $ext
				}
			}

			# Combine: base + custom (sorted) + defaults
			$allExtensions = $baseExtensions + $customExtensions + $defaultScriptExtensions

			# Build filter string (appears twice: display and pattern)
			$filterExtensions = "*." + ($allExtensions -join ";*.")
			$fileDialog.Filter = "Executable Files ($filterExtensions)|$filterExtensions|All Files (*.*)|*.*"

			$fileDialog.InitialDirectory = $txtMapFolder.Text
			
			if ($fileDialog.ShowDialog() -eq "OK") {
				$selectedPath = $fileDialog.FileName
				$selectedFile = [System.IO.Path]::GetFileName($selectedPath)

				# Use shared function to update form (file)
				Update-FormFromSelection -SelectedPath $selectedPath -FileName $selectedFile -txtMapFolder $txtMapFolder -txtSandboxFolderName $txtSandboxFolderName -txtScript $txtScript -lblStatus $lblStatus -btnSaveScript $btnSaveScript -chkNetworking $chkNetworking -chkSkipWinGet $chkSkipWinGet -cmbInstallPackages $cmbInstallPackages -wsbDir $wsbDir
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
			# Update script content
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
				if ($list -eq "AutoInstall") {
					[void]$cmbInstallPackages.Items.Add([char]0x2699 + " $list")  # âš™ AutoInstall
				} else {
					[void]$cmbInstallPackages.Items.Add($list)
				}
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
						if ($list -eq "AutoInstall") {
							[void]$this.Items.Add([char]0x2699 + " $list")  # âš™ AutoInstall
						} else {
							[void]$this.Items.Add($list)
						}
					}
					[void]$this.Items.Add("[Create new list...]")

					$this.SelectedItem = $currentSelection
					$tooltipPackages.SetToolTip($this, (Get-PackageListTooltip))
				} else {
					$this.SelectedIndex = 0
				}
			}
		})


		# KeyDown event for Delete key
		$cmbInstallPackages.Add_KeyDown({
			param($comboBox, $e)

			# Check if Delete key pressed
			if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Delete) {
				$selectedList = $comboBox.SelectedItem

				# Validation - prevent deletion of special items
				if ($selectedList -eq "" -or
					$selectedList -eq "[Create new list...]" -or
					$selectedList -like "*AutoInstall*") {  # Handles both "AutoInstall" and "âš™ AutoInstall"
					return
				}

				# Strip icon prefix if present (e.g., "âš™ AutoInstall" â†’ "AutoInstall")
				$listName = $selectedList -replace '^[^\w]+\s*', ''

				# Confirmation dialog using themed message dialog
				$result = Show-ThemedMessageDialog -Title "Confirm Delete" `
					-Message "Delete package list '$listName'?`n`nThis action cannot be undone." `
					-Buttons "OKCancel" `
					-Icon "Warning" `
					-ParentIcon $form.Icon

				if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
					# Delete file
					$listPath = Join-Path (Join-Path $Script:WorkingDir "wsb") "$listName.txt"
					if (Test-Path $listPath) {
						Remove-Item -Path $listPath -Force
					}

					# Update .ini file (set to 0)
					Set-SandboxConfig -Section 'Lists' -Key $listName -Value '0' -WorkingDir $Script:WorkingDir

					# Refresh dropdown
					$comboBox.Items.Clear()
					[void]$comboBox.Items.Add("")

					$lists = Get-PackageLists
					foreach ($list in $lists) {
						if ($list -eq "AutoInstall") {
							[void]$comboBox.Items.Add([char]0x2699 + " $list")
						} else {
							[void]$comboBox.Items.Add($list)
						}
					}
					[void]$comboBox.Items.Add("[Create new list...]")

					$comboBox.SelectedIndex = 0
					$lblStatus.Text = "Status: Package list '$listName' deleted"

					# Update tooltip
					$tooltipPackages.SetToolTip($comboBox, (Get-PackageListTooltip))
				}

				# Mark event as handled
				$e.Handled = $true
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

			# Strip icon prefix if present (e.g., "⚙ AutoInstall" → "AutoInstall")
			$listName = $selectedList -replace '^[^\w]+\s*', ''

			$result = Show-PackageListEditor -ListName $listName

			if ($result.DialogResult -eq 'OK') {
				# Reconstruct selection with icon if AutoInstall
				$currentSelection = if ($listName -eq "AutoInstall") { [char]0x2699 + " $listName" } else { $listName }
				$cmbInstallPackages.Items.Clear()
				[void]$cmbInstallPackages.Items.Add("")

				$lists = Get-PackageLists
				foreach ($list in $lists) {
					if ($list -eq "AutoInstall") {
						[void]$cmbInstallPackages.Items.Add([char]0x2699 + " $list")  # âš™ AutoInstall
					} else {
						[void]$cmbInstallPackages.Items.Add($list)
					}
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
		$y += $labelHeight + 5


		# Skip WinGet Installation checkbox
		$chkSkipWinGet = New-Object System.Windows.Forms.CheckBox
		$chkSkipWinGet.Location = New-Object System.Drawing.Point($leftMargin, $y)
		$chkSkipWinGet.Size = New-Object System.Drawing.Size($controlWidth, $controlHeight)
		$chkSkipWinGet.Text = "Skip WinGet installation (network only mode)"
		$chkSkipWinGet.Checked = $false
		$chkSkipWinGet.Enabled = $true  # Enabled by default since Networking is checked by default
		$tooltipSkipWinGet = New-Object System.Windows.Forms.ToolTip
		$tooltipSkipWinGet.SetToolTip($chkSkipWinGet, "Enable networking without installing WinGet. Useful for quick browser tests or manual downloads. Pre-install shortcuts and settings still apply.")
		$form.Controls.Add($chkSkipWinGet)

	# Add event handler for Skip WinGet checkbox
	$chkSkipWinGet.Add_CheckedChanged({
		$networkingEnabled = $chkNetworking.Checked
		$skipWinGet = $this.Checked

		# Only process if networking is enabled
		if ($networkingEnabled) {
			$winGetFeaturesEnabled = -not $skipWinGet

			# Control WinGet-related UI elements
			$cmbInstallPackages.Enabled = $winGetFeaturesEnabled
			$btnEditPackages.Enabled = $winGetFeaturesEnabled -and ($cmbInstallPackages.SelectedIndex -gt 0)
			$cmbWinGetVersion.Enabled = $winGetFeaturesEnabled -and -not $chkPrerelease.Checked
			$chkPrerelease.Enabled = $winGetFeaturesEnabled
			$chkClean.Enabled = $winGetFeaturesEnabled

			# Clear selections when enabling skip mode
			if ($skipWinGet) {
				$cmbInstallPackages.SelectedIndex = 0
				$cmbWinGetVersion.SelectedIndex = 0
				$chkPrerelease.Checked = $false
				$chkClean.Checked = $false
			}

			# Re-evaluate file selection when user changes settings
			if (-not [string]::IsNullOrWhiteSpace($script:currentSelectedFile)) {
				# Re-run package auto-selection logic
				$selectedFile = $script:currentSelectedFile

				# Auto-select package lists for specific file types (.py, .ahk, .au3)
				Update-PackageSelectionForFileType -FileName $selectedFile `
					-NetworkingCheckbox $chkNetworking `
					-SkipWinGetCheckbox $chkSkipWinGet `
					-PackageComboBox $cmbInstallPackages `
					-StatusLabel $lblStatus `
					-WorkingDir $wsbDir
			}
		}
	})

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
		$chkNetworking.Size = New-Object System.Drawing.Size(130, $labelHeight)
		$chkNetworking.Text = "Enable Networking"
		$chkNetworking.Checked = $true
		$tooltipNetworking = New-Object System.Windows.Forms.ToolTip
		$tooltipNetworking.SetToolTip($chkNetworking, "Enable network access in sandbox (required for WinGet downloads)")

		# Add event handler to enable/disable WinGet-related controls based on networking
		$chkNetworking.Add_CheckedChanged({
		$networkingEnabled = $this.Checked
		$skipWinGet = $chkSkipWinGet.Checked

		# Enable/disable "Skip WinGet" checkbox based on networking state
		$chkSkipWinGet.Enabled = $networkingEnabled

		# When networking is disabled, uncheck "Skip WinGet"
		if (-not $networkingEnabled) {
			$chkSkipWinGet.Checked = $false
		}

		# Determine if WinGet features should be enabled
		# They're only enabled when networking is ON and skip is OFF
		$winGetFeaturesEnabled = $networkingEnabled -and -not $skipWinGet

		# Control WinGet-related UI elements
		$cmbInstallPackages.Enabled = $winGetFeaturesEnabled
		# Edit button requires both networking enabled AND a valid package list selected
		$selectedItem = $cmbInstallPackages.SelectedItem
		$btnEditPackages.Enabled = $winGetFeaturesEnabled -and ($selectedItem -ne "" -and $selectedItem -ne "[Create new list...]")
		$cmbWinGetVersion.Enabled = $winGetFeaturesEnabled -and -not $chkPrerelease.Checked
		$chkPrerelease.Enabled = $winGetFeaturesEnabled
		$chkClean.Enabled = $winGetFeaturesEnabled

		# Clear selections when disabling
		if (-not $networkingEnabled) {
			$cmbInstallPackages.SelectedIndex = 0  # Select empty option
			$cmbWinGetVersion.SelectedIndex = 0    # Select empty option
			$chkPrerelease.Checked = $false
			$chkClean.Checked = $false
		}

		# Re-evaluate file selection when user changes settings
		if (-not [string]::IsNullOrWhiteSpace($script:currentSelectedFile)) {
			# Re-run package auto-selection logic
			$selectedFile = $script:currentSelectedFile

			# Auto-select package lists for specific file types (.py, .ahk, .au3)
			Update-PackageSelectionForFileType -FileName $selectedFile `
				-NetworkingCheckbox $chkNetworking `
				-SkipWinGetCheckbox $chkSkipWinGet `
				-PackageComboBox $cmbInstallPackages `
				-StatusLabel $lblStatus `
				-WorkingDir $wsbDir
		}
		})

		$form.Controls.Add($chkNetworking)

		# Protected Client checkbox (same Y-position as Networking, positioned to the right)
		$chkProtectedClient = New-Object System.Windows.Forms.CheckBox
		$chkProtectedClient.Location = New-Object System.Drawing.Point(($leftMargin + 140), $y)
		$chkProtectedClient.Size = New-Object System.Drawing.Size(200, $labelHeight)
		$chkProtectedClient.Text = "Protected Client"
		$chkProtectedClient.Checked = $false
		$form.Controls.Add($chkProtectedClient)
		$toolTip.SetToolTip($chkProtectedClient, "AppContainer Isolation mode. Provides extra security boundaries but restricts copy/paste of files")

		$y += $labelHeight + 5


		# Memory dropdown - Dynamic detection on first dropdown click (Issue #16)
		# Detection is deferred until user clicks the dropdown to avoid startup delay
		$script:memoryDetected = $false
		$script:tooltipMemory = New-Object System.Windows.Forms.ToolTip

		$lblMemory = New-Object System.Windows.Forms.Label
		$lblMemory.Location = New-Object System.Drawing.Point($leftMargin, $y)
		$lblMemory.Size = New-Object System.Drawing.Size(120, $labelHeight)
		$lblMemory.Text = "Memory (MB):"
		$form.Controls.Add($lblMemory)

		$cmbMemory = New-Object System.Windows.Forms.ComboBox
		$cmbMemory.Location = New-Object System.Drawing.Point(($leftMargin + 139), $y)
		$cmbMemory.Size = New-Object System.Drawing.Size(120, $controlHeight)
		$cmbMemory.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

		# Initialize with default value - detection happens on first dropdown click
		[void]$cmbMemory.Items.Add("4096")
		$cmbMemory.SelectedIndex = 0
		$script:tooltipMemory.SetToolTip($cmbMemory, "Click dropdown to detect available memory options")

		# Event handler for lazy memory detection on first dropdown click
		$cmbMemory.add_DropDown({
			if (-not $script:memoryDetected) {
				$script:memoryDetected = $true

				# Remember current selection
				$previousSelection = $cmbMemory.SelectedItem

				# Detect system memory using multiple methods (no elevation required):
				# 1. ComputerInfo (Win10+, preferred - fastest and no CIM)
				# 2. Win32_ComputerSystem via Get-CimInstance (fallback)
				# 3. WMI via Get-WmiObject (older systems)
				# 4. Hard-coded fallback (8 GB)
				$totalMemoryMB = $null
				$maxSafeMemory = $null

				try {
					# Method 1: Try ComputerInfo (Windows 10+ preferred method - no CIM/WMI)
					try {
						$prevProgressPreference = $ProgressPreference
						$ProgressPreference = 'SilentlyContinue'
						$computerInfo = Get-ComputerInfo -Property CsTotalPhysicalMemory -ErrorAction Stop
						$ProgressPreference = $prevProgressPreference
						$totalMemoryMB = [int]($computerInfo.CsTotalPhysicalMemory / 1MB)
						Write-Verbose "Memory detected via ComputerInfo: $totalMemoryMB MB"
					}
					catch {
						if ($prevProgressPreference) { $ProgressPreference = $prevProgressPreference }
						# Method 2: Try CIM (works without elevation on most systems)
						try {
							$totalMemoryMB = [int]((Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop).TotalPhysicalMemory / 1MB)
							Write-Verbose "Memory detected via CIM: $totalMemoryMB MB"
						}
						catch {
							# Method 3: Try WMI as last resort
							try {
								$totalMemoryMB = [int]((Get-WmiObject -Class Win32_ComputerSystem -ErrorAction Stop).TotalPhysicalMemory / 1MB)
								Write-Verbose "Memory detected via WMI: $totalMemoryMB MB"
							}
							catch {
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
					$totalMemoryMB = 8192
					$maxSafeMemory = 6144
					Write-Verbose "Could not detect system memory, using fallback: $totalMemoryMB MB (max safe: $maxSafeMemory MB)"
				}

				# Generate memory options dynamically based on available RAM
				$allMemoryOptions = @(2048, 4096, 6144, 8192, 10240, 12288, 16384, 20480, 24576, 32768, 49152, 65536)
				$memoryOptions = $allMemoryOptions | Where-Object { $_ -le $maxSafeMemory }

				# Ensure minimum option exists (2048 MB required by Windows Sandbox)
				if (-not $memoryOptions -or $memoryOptions.Count -eq 0) {
					$memoryOptions = @(2048)
				}

				# Clear and repopulate with detected options
				$cmbMemory.Items.Clear()
				$defaultIndex = -1
				$index = 0
				foreach ($memOption in $memoryOptions) {
					[void]$cmbMemory.Items.Add($memOption.ToString())
					if ($memOption -eq 4096) { $defaultIndex = $index }
					$index++
				}

				# Restore previous selection if still valid, otherwise use default
				$previousIndex = $cmbMemory.Items.IndexOf($previousSelection)
				if ($previousIndex -ge 0) {
					$cmbMemory.SelectedIndex = $previousIndex
				}
				elseif ($defaultIndex -ge 0) {
					$cmbMemory.SelectedIndex = $defaultIndex
				}
				elseif ($cmbMemory.Items.Count -gt 0) {
					$cmbMemory.SelectedIndex = $cmbMemory.Items.Count - 1
				}

				# Update tooltip with detected memory info
				$highestOption = if ($memoryOptions.Count -gt 0) { $memoryOptions[-1] } else { 2048 }
				$highestOptionGB = [math]::Round($highestOption / 1024, 1)
				$totalGB = [math]::Round($totalMemoryMB / 1024, 1)
				$script:tooltipMemory.SetToolTip($cmbMemory, "RAM for sandbox. Your system: $totalGB GB total. Highest safe option: $highestOptionGB GB (leaves 25% for Windows)")
			}
		})

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

		# Edit config button (gear icon)
		$btnEditConfig = New-Object System.Windows.Forms.Button
		$btnEditConfig.Location = New-Object System.Drawing.Point((20+25), ($y + 3))
		$btnEditConfig.Size = New-Object System.Drawing.Size(20, 20)
		$btnEditConfig.Text = [char]0x2699
		$btnEditConfig.Font = New-Object System.Drawing.Font("Segoe UI Symbol", 14, [System.Drawing.FontStyle]::Regular)

		# Tooltip
		$tooltipEditMappings = New-Object System.Windows.Forms.ToolTip
		$tooltipEditMappings.SetToolTip($btnEditConfig, "Edit config...")

		# Click event
		$btnEditConfig.Add_Click({
			Show-PackageListEditor -EditorMode "ConfigEdit" | Out-Null
		})

		$form.Controls.Add($btnEditConfig)

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
				# Reset original content for default scripts
				$script:originalScriptContent = $null
			} else {
				$txtScript.Text = ""
				$script:originalScriptContent = $null
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
				# Reset original content for default scripts
				$script:originalScriptContent = $null
			} else {
				$txtScript.Text = ""
				$script:originalScriptContent = $null
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

				# Store original content for change tracking
				$script:originalScriptContent = $txtScript.Text

				# Update Save button state for loaded file
				# Enable Save button if:
				# 1. File has CUSTOM OVERRIDE header (custom script - always editable)
				# 2. File is NOT a default script (Std-*.ps1 without custom header)
				$isCustomOverride = $scriptContent -match '^\s*#\s*CUSTOM\s+OVERRIDE'
				$isDefaultScript = Test-IsDefaultScript -FilePath $openFileDialog.FileName

				if ($isCustomOverride -or -not $isDefaultScript) {
					$btnSaveScript.Enabled = $true
				} else {
					$btnSaveScript.Enabled = $false
				}

				# Update status to show loaded script
				$scriptFileName = [System.IO.Path]::GetFileName($openFileDialog.FileName)
				if ($isCustomOverride) {
					$lblStatus.Text = "Status: Loaded $scriptFileName (CUSTOM)"
				} else {
					$lblStatus.Text = "Status: Loaded $scriptFileName"
				}
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
				Show-ThemedMessageDialog -Title "Save Error" -Message "No script content to save." -Buttons "OK" -Icon "Warning" -ParentIcon $form.Icon | Out-Null
				return
			}

			# If we have a current file, save directly (silent save)
			if ($script:currentScriptFile -and (Test-Path (Split-Path $script:currentScriptFile -Parent))) {
				try {
					$txtScript.Text | Out-File -FilePath $script:currentScriptFile -Encoding UTF8
					# Update original content and disable Save button after successful save
					$script:originalScriptContent = $txtScript.Text
					$btnSaveScript.Enabled = $false
					# Silent save - no success dialog
				}
				catch {
					Show-ThemedMessageDialog -Title "Save Error" -Message "Error saving script: $($_.Exception.Message)" -Buttons "OK" -Icon "Error" -ParentIcon $form.Icon | Out-Null
				}
			}
			else {
				# No current file - show themed input dialog for filename
				$scriptName = Show-ThemedInputDialog -Title "Save Script" -Prompt "Enter script filename (without .ps1):" -DefaultValue "" -ParentIcon $form.Icon

				if (-not [string]::IsNullOrWhiteSpace($scriptName)) {
					# Auto-append .ps1 extension if missing
					$scriptNameStr = [string]$scriptName
					$fileName = if ($scriptNameStr.EndsWith('.ps1')) { $scriptNameStr } else { "$scriptNameStr.ps1" }
					$targetPath = Join-Path $wsbDir $fileName

					# Check if file exists - show overwrite confirmation
					if (Test-Path $targetPath) {
						$result = Show-ThemedMessageDialog -Title "File Exists" -Message "File already exists:`n$targetPath`n`nOverwrite?" -Buttons "OKCancel" -Icon "Warning" -ParentIcon $form.Icon
						if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
							return  # User canceled
						}
					}

					# Save file
					try {
						# Ensure wsb directory exists
						if (-not (Test-Path $wsbDir)) {
							New-Item -ItemType Directory -Path $wsbDir -Force | Out-Null
						}
						$txtScript.Text | Out-File -FilePath $targetPath -Encoding UTF8

						# Update current file tracking and original content
						$script:currentScriptFile = $targetPath
						$script:originalScriptContent = $txtScript.Text
						$btnSaveScript.Enabled = $false

						Show-ThemedMessageDialog -Title "Save Complete" -Message "Script saved successfully to:`n$targetPath" -Buttons "OK" -Icon "Information" -ParentIcon $form.Icon | Out-Null
					}
					catch {
						Show-ThemedMessageDialog -Title "Save Error" -Message "Error saving script: $($_.Exception.Message)" -Buttons "OK" -Icon "Error" -ParentIcon $form.Icon | Out-Null
					}
				}
			}
		})
		$form.Controls.Add($btnSaveScript)

		# Set initial Save button state based on current script
		if ([string]::IsNullOrWhiteSpace($txtScript.Text)) {
			$btnSaveScript.Enabled = $false
		} elseif ($script:currentScriptFile) {
			# Check if script has CUSTOM OVERRIDE header
			$hasCustomOverride = $txtScript.Text -match '^\s*#\s*CUSTOM\s+OVERRIDE'
			$isDefaultScript = $false
			if (-not $hasCustomOverride) {
				$isDefaultScript = Test-IsDefaultScript -FilePath $script:currentScriptFile
			}
			$btnSaveScript.Enabled = -not $isDefaultScript
		}

		# Add TextChanged event to update Save button state dynamically
		$txtScript.Add_TextChanged({
		# Check if current script has CUSTOM OVERRIDE header
		$currentContent = $txtScript.Text
		$hasCustomOverride = $currentContent -match '^\s*#\s*CUSTOM\s+OVERRIDE'

		# Check if file is a default script (but allow custom override to bypass this)
		$isDefaultScript = $false
		if (-not $hasCustomOverride) {
			$isDefaultScript = Test-IsDefaultScript -FilePath $script:currentScriptFile
		}

		if ($isDefaultScript) {
			# Default scripts without CUSTOM OVERRIDE cannot be saved
			$btnSaveScript.Enabled = $false
		} elseif ([string]::IsNullOrWhiteSpace($txtScript.Text)) {
			# Empty script cannot be saved
			$btnSaveScript.Enabled = $false
		} elseif ([string]::IsNullOrWhiteSpace($script:currentScriptFile)) {
			# No file path - must use Save As
			$btnSaveScript.Enabled = $false
		} else {
			# Enable Save button only if content has changed
			$hasChanged = ($null -eq $script:originalScriptContent) -or ($currentContent -ne $script:originalScriptContent)
			$btnSaveScript.Enabled = $hasChanged
		}
	})


		$btnSaveAsScript = New-Object System.Windows.Forms.Button
		$btnSaveAsScript.Location = New-Object System.Drawing.Point(263, $y)
		$btnSaveAsScript.Size = New-Object System.Drawing.Size(75, $controlHeight)
		$btnSaveAsScript.Text = "Save as..."
		$btnSaveAsScript.Add_Click({
			if ([string]::IsNullOrWhiteSpace($txtScript.Text)) {
				Show-ThemedMessageDialog -Title "Save Error" -Message "No script content to save." -Buttons "OK" -Icon "Warning" -ParentIcon $form.Icon | Out-Null
				return
			}

			# Get default filename (if current script is custom, not default)
			$defaultName = ""
			if ($script:currentScriptFile -and -not (Test-IsDefaultScript -FilePath $script:currentScriptFile)) {
				$defaultName = [System.IO.Path]::GetFileNameWithoutExtension($script:currentScriptFile)
			}

			# Show themed input dialog for filename
			$scriptName = Show-ThemedInputDialog -Title "Save Script As" -Prompt "Enter script filename (without .ps1):" -DefaultValue $defaultName -ParentIcon $form.Icon

			if (-not [string]::IsNullOrWhiteSpace($scriptName)) {
				# Auto-append .ps1 extension if missing
				$scriptNameStr = [string]$scriptName
				$fileName = if ($scriptNameStr.EndsWith('.ps1')) { $scriptNameStr } else { "$scriptNameStr.ps1" }
				$targetPath = Join-Path $wsbDir $fileName

				# Check if file exists - show overwrite confirmation
				if (Test-Path $targetPath) {
					$result = Show-ThemedMessageDialog -Title "File Exists" -Message "File already exists:`n$targetPath`n`nOverwrite?" -Buttons "OKCancel" -Icon "Warning" -ParentIcon $form.Icon
					if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
						return  # User canceled
					}
				}

				# Save file
				try {
					# Ensure wsb directory exists
					if (-not (Test-Path $wsbDir)) {
						New-Item -ItemType Directory -Path $wsbDir -Force | Out-Null
					}
					$txtScript.Text | Out-File -FilePath $targetPath -Encoding UTF8

					# Update current file tracking and original content
					$script:currentScriptFile = $targetPath
					$script:originalScriptContent = $txtScript.Text
					$btnSaveScript.Enabled = $false

					Show-ThemedMessageDialog -Title "Save Complete" -Message "Script saved successfully to:`n$targetPath" -Buttons "OK" -Icon "Information" -ParentIcon $form.Icon | Out-Null
				}
				catch {
					Show-ThemedMessageDialog -Title "Save Error" -Message "Error saving script: $($_.Exception.Message)" -Buttons "OK" -Icon "Error" -ParentIcon $form.Icon | Out-Null
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
			SkipWinGetInstallation = $chkSkipWinGet.Checked
			MapFolderReadOnly = $chkMapFolderReadOnly.Checked
				MemoryInMB = [int]$cmbMemory.SelectedItem
				vGPU = $cmbvGPU.SelectedItem
				ProtectedClient = $chkProtectedClient.Checked
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

		# Enable AutoScroll for high DPI scaling scenarios (issue #4)
		# Must be done in Form.Load event when ClientSize is finalized
		$finalY = $y  # Capture final Y position before Form.Load
		$form.Add_Load({
			# Calculate actual content height
			$contentHeight = $finalY + 50  # Button Y + button height + margin
			$clientHeight = $this.ClientSize.Height

			# Get screen working area (excludes taskbar)
			$workingArea = [System.Windows.Forms.Screen]::FromControl($this).WorkingArea
			$screenHeight = $workingArea.Height
			$screenTop = $workingArea.Top

			# Calculate if form exceeds screen working area
			$formTop = $this.Top
			$formBottom = $formTop + $this.Height
			$formExceedsScreen = $formBottom -gt $screenHeight

			# If form exceeds screen working area, resize it and enable scrolling
			if ($formExceedsScreen) {
				# Calculate maximum form height that fits in working area (with small margin)
				$maxFormHeight = $screenHeight - 20  # 20px margin from taskbar

				# Resize form to fit screen
				$this.Height = $maxFormHeight

				# Reposition to top of working area
				$this.Top = $screenTop + 10  # 10px margin from top

				# Enable scrolling since we had to shrink the form
				$this.AutoScroll = $true
				# Set AutoScrollMinSize - subtract scrollbar width from client width to prevent horizontal scrollbar
				$scrollBarWidth = [System.Windows.Forms.SystemInformation]::VerticalScrollBarWidth
				$this.AutoScrollMinSize = New-Object System.Drawing.Size(($this.ClientSize.Width - $scrollBarWidth - 5), $contentHeight)
				# Explicitly disable horizontal scrolling
				$this.HorizontalScroll.Enabled = $false
				$this.HorizontalScroll.Visible = $false
				$this.HorizontalScroll.Maximum = 0
			}
			# Enable scrolling if content exceeds client height (even if form fits on screen)
			elseif ($contentHeight -gt $clientHeight) {
				$this.AutoScroll = $true
				# Set AutoScrollMinSize - subtract scrollbar width from client width to prevent horizontal scrollbar
				$scrollBarWidth = [System.Windows.Forms.SystemInformation]::VerticalScrollBarWidth
				$this.AutoScrollMinSize = New-Object System.Drawing.Size(($this.ClientSize.Width - $scrollBarWidth - 5), $contentHeight)
				# Explicitly disable horizontal scrolling
				$this.HorizontalScroll.Enabled = $false
				$this.HorizontalScroll.Visible = $false
				$this.HorizontalScroll.Maximum = 0
			}
		})

		# Add Shown event to force hide horizontal scrollbar after form is fully rendered
		$form.Add_Shown({
			if ($this.AutoScroll) {
				# Force hide horizontal scrollbar after form is shown
				$this.HorizontalScroll.Enabled = $false
				$this.HorizontalScroll.Visible = $false
				$this.HorizontalScroll.Maximum = 0
				$this.PerformLayout()  # Force layout refresh
			}

			# Set initial paths from context menu parameters (and process them)
			if ($script:InitialFolderPath -and (Test-Path $script:InitialFolderPath)) {
				# Folder selected from context menu
				Update-FormFromSelection -SelectedPath $script:InitialFolderPath -txtMapFolder $txtMapFolder -txtSandboxFolderName $txtSandboxFolderName -txtScript $txtScript -lblStatus $lblStatus -btnSaveScript $btnSaveScript -chkNetworking $chkNetworking -chkSkipWinGet $chkSkipWinGet -cmbInstallPackages $cmbInstallPackages -wsbDir $wsbDir
		}
			if ($script:InitialFilePath -and (Test-Path $script:InitialFilePath)) {
				# File selected from context menu
				$selectedFile = [System.IO.Path]::GetFileName($script:InitialFilePath)
				Update-FormFromSelection -SelectedPath $script:InitialFilePath -FileName $selectedFile -txtMapFolder $txtMapFolder -txtSandboxFolderName $txtSandboxFolderName -txtScript $txtScript -lblStatus $lblStatus -btnSaveScript $btnSaveScript -chkNetworking $chkNetworking -chkSkipWinGet $chkSkipWinGet -cmbInstallPackages $cmbInstallPackages -wsbDir $wsbDir
		}



		})

		# Apply theme based on saved preference (BEFORE context menu to ensure colors are set)
		Set-ThemeToForm -Form $form -UpdateButtonColor $updateButtonGreen

		# Prepare shell integration scriptblocks (only for SandboxStart, not WAU-Settings-GUI)
		$sandboxStartScript = Join-Path $Script:WorkingDir 'SandboxStart.ps1'
		$shellIntegrationParams = @{}

		if (Test-Path $sandboxStartScript) {
			# Capture WorkingDir in local variable for use in scriptblocks
			$localWorkingDir = $Script:WorkingDir

			# Helper function to test registry key existence using reg.exe
			$testRegKey = {
				param([string]$KeyPath)
				$psi = New-Object System.Diagnostics.ProcessStartInfo
				$psi.FileName = 'reg.exe'
				$psi.Arguments = "query `"$KeyPath`""
				$psi.RedirectStandardOutput = $true
				$psi.RedirectStandardError = $true
				$psi.UseShellExecute = $false
				$psi.CreateNoWindow = $true
				$p = [System.Diagnostics.Process]::Start($psi)
				$output = $p.StandardOutput.ReadToEnd()
				$p.WaitForExit()
				return ($output -match 'SandboxStart')
			}

			# Test if context menu integration is installed
			$testContextMenu = {
				$folderKeyReg = 'HKCU\Software\Classes\Directory\shell\SandboxStart'
				$fileKeyReg = 'HKCU\Software\Classes\*\shell\SandboxStart'
				$driveKeyReg = 'HKCU\Software\Classes\Drive\shell\SandboxStart'

				$folderExists = & $testRegKey $folderKeyReg
				$fileExists = & $testRegKey $fileKeyReg
				$driveExists = & $testRegKey $driveKeyReg

				return ($folderExists -and $fileExists -and $driveExists)
			}

			# Update context menu integration
			$updateContextMenu = {
				param([string]$WorkingDir, [bool]$Remove)

				try {
					Write-Verbose "updateContextMenu called: WorkingDir=$WorkingDir, Remove=$Remove"

					$scriptPath = Join-Path $WorkingDir 'SandboxStart.ps1'
					$iconPath = Join-Path $WorkingDir 'startmenu-icon.ico'
					$folderKey = 'HKCU:\Software\Classes\Directory\shell\SandboxStart'
					$folderKeyReg = 'HKCU\Software\Classes\Directory\shell\SandboxStart'
					$fileKeyReg = 'HKCU\Software\Classes\*\shell\SandboxStart'
					$driveKey = 'HKCU:\Software\Classes\Drive\shell\SandboxStart'
					$driveKeyReg = 'HKCU\Software\Classes\Drive\shell\SandboxStart'

					if ($Remove) {
						Write-Verbose "Removing context menu entries..."

						# Remove folder context menu (can use Remove-Item for non-* paths)
						if (& $testRegKey $folderKeyReg) {
							Remove-Item $folderKey -Recurse -Force -ErrorAction Stop
							Write-Verbose "Removed folder key"
						}

						# Remove file context menu (use reg.exe delete for * path to avoid hang)
						if (& $testRegKey $fileKeyReg) {
							$null = reg.exe delete "$fileKeyReg" /f 2>&1
							Write-Verbose "Removed file key"
						}

						# Remove drive context menu (can use Remove-Item for non-* paths)
						if (& $testRegKey $driveKeyReg) {
							Remove-Item $driveKey -Recurse -Force -ErrorAction Stop
							Write-Verbose "Removed drive key"
						}

						return $true
					}

					# Create folder context menu
					Write-Verbose "Creating folder context menu..."
					if (-not (Test-Path $folderKey)) {
						Write-Verbose "Creating new folder key: $folderKey"
						$null = New-Item -Path $folderKey -Force -ErrorAction Stop
					}
					$null = New-ItemProperty -Path $folderKey -Name '(Default)' -Value 'Test in Windows Sandbox' -PropertyType String -Force -ErrorAction Stop
					$null = New-ItemProperty -Path $folderKey -Name 'Icon' -Value $iconPath -PropertyType String -Force -ErrorAction Stop

					$folderCommandKey = "$folderKey\command"
					if (-not (Test-Path $folderCommandKey)) {
						$null = New-Item -Path $folderCommandKey -Force -ErrorAction Stop
					}
					$null = New-ItemProperty -Path $folderCommandKey -Name '(Default)' -Value "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -FolderPath `"%V`"" -PropertyType String -Force -ErrorAction Stop
					Write-Verbose "Folder context menu created"

					# Create file context menu using reg.exe (PowerShell cmdlets hang on * path)
					Write-Verbose "Creating file context menu..."
					$fileCommandKeyReg = "$fileKeyReg\command"
					$fileCommand = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File \`"$scriptPath\`" -FilePath \`"%1\`""

					# Create all registry entries in batch
					$null = reg.exe add "$fileKeyReg" /ve /d "Test in Windows Sandbox" /f 2>&1
					$null = reg.exe add "$fileKeyReg" /v Icon /d "$iconPath" /f 2>&1
					$null = reg.exe add "$fileCommandKeyReg" /ve /d "$fileCommand" /f 2>&1
					Write-Verbose "File context menu created"

					# Create drive context menu (same pattern as folder)
					Write-Verbose "Creating drive context menu..."
					if (-not (Test-Path $driveKey)) {
						Write-Verbose "Creating new drive key: $driveKey"
						$null = New-Item -Path $driveKey -Force -ErrorAction Stop
					}
					$null = New-ItemProperty -Path $driveKey -Name '(Default)' -Value 'Test in Windows Sandbox' -PropertyType String -Force -ErrorAction Stop
					$null = New-ItemProperty -Path $driveKey -Name 'Icon' -Value $iconPath -PropertyType String -Force -ErrorAction Stop

					$driveCommandKey = "$driveKey\command"
					if (-not (Test-Path $driveCommandKey)) {
						$null = New-Item -Path $driveCommandKey -Force -ErrorAction Stop
					}
					$null = New-ItemProperty -Path $driveCommandKey -Name '(Default)' -Value "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -FolderPath `"%V\`"" -PropertyType String -Force -ErrorAction Stop
					Write-Verbose "Drive context menu created"

					return $true
				}
				catch {
					$errorMsg = "updateContextMenu failed: $_`nStack Trace: $($_.ScriptStackTrace)"
					Write-Error $errorMsg
					[System.Windows.Forms.MessageBox]::Show($errorMsg, 'Context Menu Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
					return $false
				}
			}

			$shellIntegrationParams = @{
				WorkingDir = $localWorkingDir
				TestRegKey = $testRegKey
				TestContextMenu = $testContextMenu
				UpdateContextMenu = $updateContextMenu
				AppIcon = $appIcon
			}
		}

		# Attach right-click context menu for theme selection (AFTER theme is applied)
		$contextMenu = Show-ThemeContextMenu -Form $form -UpdateButtonColor $updateButtonGreen @shellIntegrationParams
		$form.ContextMenuStrip = $contextMenu

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
if ($dialogResult.SkipWinGetInstallation) {
	$sandboxParams.SkipWinGetInstallation = $dialogResult.SkipWinGetInstallation
}
if ($dialogResult.MemoryInMB) {
	$sandboxParams.MemoryInMB = $dialogResult.MemoryInMB
}
if (![string]::IsNullOrWhiteSpace($dialogResult.vGPU)) {
	$sandboxParams.vGPU = $dialogResult.vGPU
}
if ($dialogResult.ProtectedClient) {
	$sandboxParams.ProtectedClient = $dialogResult.ProtectedClient
}
if ($dialogResult.MapFolderReadOnly) {
	$sandboxParams.MapFolderReadOnly = $dialogResult.MapFolderReadOnly
}

# Call SandboxTest with collected parameters
SandboxTest @sandboxParams

# Wait for key press if requested
if ($dialogResult.Wait) {
	Write-Host "`nPress any key to exit..." -ForegroundColor Yellow
	$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

exit


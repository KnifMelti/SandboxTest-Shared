#Requires -Version 5.1

<#
.SYNOPSIS
Common utility functions for SandboxStart scripts

.DESCRIPTION
Shared helper functions to eliminate code duplication across
Show-SandboxTestDialog.ps1 and SandboxTest.ps1

.NOTES
Author: SandboxStart Project
#>

function Ensure-DirectoryExists {
	<#
	.SYNOPSIS
	Creates a directory if it doesn't exist

	.PARAMETER Path
	The directory path to ensure exists

	.OUTPUTS
	Returns the path (for chaining)
	#>
	param(
		[Parameter(Mandatory)]
		[string]$Path
	)

	if (-not (Test-Path $Path)) {
		New-Item -ItemType Directory -Path $Path -Force | Out-Null
	}
	return $Path
}

function Write-AsciiFile {
	<#
	.SYNOPSIS
	Writes content to a file with ASCII encoding

	.PARAMETER Path
	The file path to write to

	.PARAMETER Content
	The content to write

	.PARAMETER Force
	Overwrite existing file if present
	#>
	param(
		[Parameter(Mandatory)]
		[string]$Path,

		[Parameter(Mandatory)]
		[AllowEmptyString()]
		[string]$Content,

		[switch]$Force
	)

	$Content | Out-File -FilePath $Path -Encoding ASCII -Force:$Force
}

function Join-PathMulti {
	<#
	.SYNOPSIS
	Joins multiple path segments into a single path

	.PARAMETER Paths
	Array of path segments to join

	.EXAMPLE
	Join-PathMulti @($env:LOCALAPPDATA, 'Packages', 'MyApp')
	#>
	param(
		[Parameter(Mandatory)]
		[string[]]$Paths
	)

	if ($Paths.Length -eq 0) {
		throw "At least one path segment is required"
	}

	$result = $Paths[0]
	for ($i = 1; $i -lt $Paths.Length; $i++) {
		$result = Join-Path $result $Paths[$i]
	}
	return $result
}

function Test-ValidFolderName {
	<#
	.SYNOPSIS
	Tests if a string is a valid folder name (not a drive root)

	.PARAMETER Name
	The folder name to test
	#>
	param(
		[string]$Name
	)

	return (![string]::IsNullOrWhiteSpace($Name) -and $Name -notmatch ':' -and $Name -ne '\')
}

function Get-DriveLetterFromPath {
	<#
	.SYNOPSIS
	Extracts the drive letter from a path

	.PARAMETER Path
	The path to extract drive letter from

	.EXAMPLE
	Get-DriveLetterFromPath 'C:\' returns 'C'
	#>
	param(
		[Parameter(Mandatory)]
		[string]$Path
	)

	return $Path.TrimEnd('\').Replace(':', '')
}

function Get-SandboxFolderName {
	<#
	.SYNOPSIS
	Determines an appropriate folder name for sandbox mapping

	.DESCRIPTION
	Converts a path to a suitable sandbox folder name:
	- Regular folder: Uses folder name
	- Drive root (C:\): Uses "Drive_C"
	- Empty/invalid: Uses default name

	.PARAMETER Path
	The source path to convert

	.PARAMETER DefaultName
	Default name if path is invalid (default: 'MappedFolder')

	.EXAMPLE
	Get-SandboxFolderName 'C:\Projects\MyApp' returns 'MyApp'
	Get-SandboxFolderName 'D:\' returns 'Drive_D'
	#>
	param(
		[Parameter(Mandatory)]
		[string]$Path,

		[string]$DefaultName = 'MappedFolder'
	)

	$folderName = Split-Path $Path -Leaf

	# Check if it's a valid folder name (not a drive root)
	if (Test-ValidFolderName -Name $folderName) {
		return $folderName
	}

	# Try to extract drive letter for root drives
	$driveLetter = Get-DriveLetterFromPath -Path $Path
	if (![string]::IsNullOrWhiteSpace($driveLetter)) {
		return "Drive_$driveLetter"
	}

	# Fallback to default
	return $DefaultName
}

function Invoke-SilentProgress {
	<#
	.SYNOPSIS
	Executes a script block with progress output suppressed

	.PARAMETER ScriptBlock
	The script block to execute

	.EXAMPLE
	Invoke-SilentProgress { Get-ComputerInfo }
	#>
	param(
		[Parameter(Mandatory)]
		[scriptblock]$ScriptBlock
	)

	$oldPref = $ProgressPreference
	try {
		$ProgressPreference = 'SilentlyContinue'
		& $ScriptBlock
	}
	finally {
		$ProgressPreference = $oldPref
	}
}

function Read-FileContent {
	<#
	.SYNOPSIS
	Reads entire file content as a single string

	.PARAMETER Path
	The file path to read

	.PARAMETER ErrorAction
	Error action preference (default: SilentlyContinue)
	#>
	param(
		[Parameter(Mandatory)]
		[string]$Path,

		[string]$ErrorAction = 'SilentlyContinue'
	)

	return (Get-Content -Path $Path -Raw -ErrorAction $ErrorAction)
}

# Export all functions
Export-ModuleMember -Function @(
	'Ensure-DirectoryExists',
	'Write-AsciiFile',
	'Join-PathMulti',
	'Test-ValidFolderName',
	'Get-DriveLetterFromPath',
	'Get-SandboxFolderName',
	'Invoke-SilentProgress',
	'Read-FileContent'
)

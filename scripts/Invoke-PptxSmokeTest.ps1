param(
	[Parameter(Mandatory = $true)]
	[string]$InputMarkdown,

	[string]$OutputPath,
	[string]$ThemeCss,
	[string]$Template,
	[ValidateSet("Debug", "Release")]
	[string]$Configuration = "Debug",
	[switch]$CiSafe,
	[switch]$SkipPowerPoint,
	[switch]$SkipRoundTripSave,
	[string]$RoundTripCopyPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Import-OpenXmlValidationAssemblies {
	param(
		[Parameter(Mandatory = $true)]
		[string]$CliOutputDirectory
	)

	$assemblyPaths = @(
		"System.IO.Packaging.dll",
		"DocumentFormat.OpenXml.Framework.dll",
		"DocumentFormat.OpenXml.dll"
	) | ForEach-Object {
		Join-Path $CliOutputDirectory $_
	}

	foreach ($assemblyPath in $assemblyPaths) {
		if (-not (Test-Path $assemblyPath)) {
			throw "Required validation assembly was not found: $assemblyPath"
		}

		if (-not ([AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.Location -eq $assemblyPath })) {
			Add-Type -Path $assemblyPath
		}
	}
}

function Test-OpenXmlPackage {
	param(
		[Parameter(Mandatory = $true)]
		[string]$PptxPath
	)

	$validationErrors = @()
	$document = $null

	try {
		$document = [DocumentFormat.OpenXml.Packaging.PresentationDocument]::Open($PptxPath, $false)
		$validator = [DocumentFormat.OpenXml.Validation.OpenXmlValidator]::new()
		$validationErrors = @($validator.Validate($document))
	}
	finally {
		if ($document) {
			$document.Dispose()
		}
	}

	if ($validationErrors.Count -gt 0) {
		Write-Host "Open XML validation failed:" -ForegroundColor Red
		foreach ($validationError in $validationErrors) {
			Write-Host ("- {0} at {1}" -f $validationError.Description, $validationError.Path.XPath)
		}

		throw "Open XML validation found $($validationErrors.Count) error(s)."
	}

	Write-Host "Open XML validation passed." -ForegroundColor Green
}

function Test-PowerPointAutomationAvailable {
	if (-not $IsWindows) {
		return $false
	}

	try {
		$powerPointType = [System.Type]::GetTypeFromProgID("PowerPoint.Application", $false)
		return $null -ne $powerPointType
	}
	catch {
		return $false
	}
}

$repoRoot = Split-Path -Parent $PSScriptRoot
$generateScript = Join-Path $PSScriptRoot "Generate-LocalPptx.ps1"
$powerPointScript = Join-Path $PSScriptRoot "Test-PowerPointOpen.ps1"

if (-not $OutputPath) {
	$inputBaseName = [System.IO.Path]::GetFileNameWithoutExtension($InputMarkdown)
	$smokeTestDirectory = Join-Path (Join-Path $repoRoot "artifacts") "smoke-tests"
	$OutputPath = Join-Path $smokeTestDirectory ("{0}-{1}.pptx" -f $inputBaseName, $Configuration.ToLowerInvariant())
}

$resolvedOutputPath = [System.IO.Path]::GetFullPath($OutputPath, $repoRoot)
$outputDirectory = Split-Path -Parent $resolvedOutputPath
if ($outputDirectory) {
	New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
}

if (-not $SkipPowerPoint -and -not $SkipRoundTripSave -and -not $RoundTripCopyPath) {
	$outputBaseName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedOutputPath)
	$RoundTripCopyPath = Join-Path $outputDirectory ("{0}-roundtrip.pptx" -f $outputBaseName)
}

$powerPointAvailable = Test-PowerPointAutomationAvailable

if ($CiSafe -and -not $SkipPowerPoint -and -not $powerPointAvailable) {
	$SkipPowerPoint = $true
	Write-Host "CI-safe mode: PowerPoint COM automation is unavailable, so the smoke test will stop after Open XML validation." -ForegroundColor Yellow
}

Write-Host "Step 1: Generate PPTX from the local workspace code." -ForegroundColor Cyan
$generateArguments = @{
	InputMarkdown = $InputMarkdown
	OutputPath = $resolvedOutputPath
	Configuration = $Configuration
}

if ($ThemeCss) {
	$generateArguments.ThemeCss = $ThemeCss
}

if ($Template) {
	$generateArguments.Template = $Template
}

& $generateScript @generateArguments

$cliOutputDirectory = Join-Path (Join-Path (Join-Path (Join-Path $repoRoot "src") "MarpToPptx.Cli") "bin") (Join-Path $Configuration "net10.0")
if (-not (Test-Path $cliOutputDirectory)) {
	throw "CLI output directory was not found after generation: $cliOutputDirectory"
}

Write-Host "Step 2: Validate the generated package with Open XML SDK." -ForegroundColor Cyan
Import-OpenXmlValidationAssemblies -CliOutputDirectory $cliOutputDirectory
Test-OpenXmlPackage -PptxPath $resolvedOutputPath

if (-not $SkipPowerPoint) {
	Write-Host "Step 3: Open the package in PowerPoint." -ForegroundColor Cyan
	$powerPointArguments = @{
		PptxPath = $resolvedOutputPath
	}

	if (-not $SkipRoundTripSave -and $RoundTripCopyPath) {
		$powerPointArguments.SaveCopyAs = [System.IO.Path]::GetFullPath($RoundTripCopyPath, $repoRoot)
	}

	& $powerPointScript @powerPointArguments
}
else {
	Write-Host "Step 3: Skipped PowerPoint validation." -ForegroundColor Yellow
}

Write-Host "Smoke test passed for '$resolvedOutputPath'." -ForegroundColor Green

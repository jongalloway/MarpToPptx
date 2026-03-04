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

function Invoke-OpenXmlValidation {
	param(
		[Parameter(Mandatory = $true)]
		[string]$PptxPath,

		[Parameter(Mandatory = $true)]
		[string]$Configuration,

		[Parameter(Mandatory = $true)]
		[string]$RepositoryRoot
	)

	Push-Location $RepositoryRoot
	try {
		$arguments = @(
			"run"
			"--project"
			"src/MarpToPptx.OpenXmlValidator"
			"-c"
			$Configuration
			"--"
			$PptxPath
		)

		Write-Host ("dotnet {0}" -f ($arguments -join " "))
		& dotnet @arguments

		if ($LASTEXITCODE -ne 0) {
			throw "Open XML validation helper failed with exit code $LASTEXITCODE"
		}
	}
	finally {
		Pop-Location
	}
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

Write-Host "Step 2: Validate the generated package with Open XML SDK." -ForegroundColor Cyan
Invoke-OpenXmlValidation -PptxPath $resolvedOutputPath -Configuration $Configuration -RepositoryRoot $repoRoot

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

param(
	[Parameter(Mandatory = $true)]
	[string]$InputMarkdown,

	[string]$OutputPath,
	[string]$ThemeCss,
	[string]$Template,
	[switch]$AllowRemoteAssets,
	[ValidateSet("Debug", "Release")]
	[string]$Configuration = "Debug",
	[switch]$CiSafe,
	[switch]$SkipPowerPoint,
	[switch]$SkipRoundTripSave,
	[string]$RoundTripCopyPath,
	[switch]$ExportSlides,
	[ValidateSet("jpg", "png")]
	[string]$SlideExportFormat = "png",
	[string]$SlideExportDirectory,
	[switch]$SkipContrastAudit,
	[string]$ContrastReportPath
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

function Invoke-ContrastAudit {
	param(
		[Parameter(Mandatory = $true)]
		[string]$PptxPath,

		[Parameter(Mandatory = $true)]
		[string]$Configuration,

		[Parameter(Mandatory = $true)]
		[string]$RepositoryRoot,

		[string]$ReportPath
	)

	Push-Location $RepositoryRoot
	try {
		$arguments = @(
			"run"
			"--project"
			"src/MarpToPptx.ContrastAuditor"
			"-c"
			$Configuration
			"--"
			$PptxPath
		)

		Write-Host ("dotnet {0}" -f ($arguments -join " "))
		$lines = & dotnet @arguments 2>&1
		$exitCode = $LASTEXITCODE

		if ($ReportPath) {
			$reportDirectory = Split-Path -Parent $ReportPath
			if ($reportDirectory) {
				New-Item -ItemType Directory -Path $reportDirectory -Force | Out-Null
			}

			$lines | Out-File -FilePath $ReportPath -Encoding utf8
			Write-Host "Wrote contrast audit report to '$ReportPath'." -ForegroundColor DarkGray
		}

		if ($lines) {
			$lines | ForEach-Object { Write-Host $_ }
		}

		if ($exitCode -ne 0) {
			throw "Contrast audit failed with exit code $exitCode"
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
$exportSlidesScript = Join-Path $PSScriptRoot "Export-PptxSlides.ps1"
$usedDefaultOutputPath = $false

if (-not $OutputPath) {
	$inputBaseName = [System.IO.Path]::GetFileNameWithoutExtension($InputMarkdown)
	$smokeTestDirectory = Join-Path (Join-Path $repoRoot "artifacts") "smoke-tests"
	$OutputPath = Join-Path $smokeTestDirectory ("{0}-generated-{1}.pptx" -f $inputBaseName, $Configuration.ToLowerInvariant())
	$usedDefaultOutputPath = $true
}

$resolvedOutputPath = [System.IO.Path]::GetFullPath($OutputPath, $repoRoot)
$outputDirectory = Split-Path -Parent $resolvedOutputPath
if ($outputDirectory) {
	New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
}

if (-not $SkipPowerPoint -and -not $SkipRoundTripSave -and -not $RoundTripCopyPath) {
	if ($usedDefaultOutputPath) {
		$inputBaseName = [System.IO.Path]::GetFileNameWithoutExtension($InputMarkdown)
		$RoundTripCopyPath = Join-Path $outputDirectory ("{0}-powerpoint-resaved-{1}.pptx" -f $inputBaseName, $Configuration.ToLowerInvariant())
	}
	else {
		$outputBaseName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedOutputPath)
		$RoundTripCopyPath = Join-Path $outputDirectory ("{0}-powerpoint-resaved.pptx" -f $outputBaseName)
	}
}

if (-not $SkipContrastAudit -and -not $ContrastReportPath) {
	$outputBaseName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedOutputPath)
	$ContrastReportPath = Join-Path $outputDirectory ("{0}-contrast-audit.txt" -f $outputBaseName)
}

$powerPointAvailable = Test-PowerPointAutomationAvailable

if ($CiSafe -and -not $SkipPowerPoint -and -not $powerPointAvailable) {
	$SkipPowerPoint = $true
	Write-Host "CI-safe mode: PowerPoint COM automation is unavailable, so the smoke test will stop after Open XML validation." -ForegroundColor Yellow
}

Write-Host "Step 1: Generate PPTX from the local workspace code." -ForegroundColor Cyan
$generateArguments = @{
	InputMarkdown = $InputMarkdown
	OutputPath    = $resolvedOutputPath
	Configuration = $Configuration
}

if ($ThemeCss) {
	$generateArguments.ThemeCss = $ThemeCss
}

if ($Template) {
	$generateArguments.Template = $Template
}

if ($AllowRemoteAssets) {
	$generateArguments.AllowRemoteAssets = $true
}

try {
	& $generateScript @generateArguments
}
catch {
	Write-Error $_.Exception.Message -ErrorAction Continue
	$global:LASTEXITCODE = 1
	return
}

Write-Host "Step 2: Validate the generated package with Open XML SDK." -ForegroundColor Cyan
Invoke-OpenXmlValidation -PptxPath $resolvedOutputPath -Configuration $Configuration -RepositoryRoot $repoRoot

if (-not $SkipContrastAudit) {
	Write-Host "Step 3: Audit rendered text contrast." -ForegroundColor Cyan
	Invoke-ContrastAudit -PptxPath $resolvedOutputPath -Configuration $Configuration -RepositoryRoot $repoRoot -ReportPath $ContrastReportPath
}
else {
	Write-Host "Step 3: Skipped contrast audit." -ForegroundColor Yellow
}

if (-not $SkipPowerPoint) {
	Write-Host "Step 4: Open the package in PowerPoint." -ForegroundColor Cyan
	$powerPointArguments = @{
		PptxPath = $resolvedOutputPath
	}

	if (-not $SkipRoundTripSave -and $RoundTripCopyPath) {
		$powerPointArguments.SaveCopyAs = [System.IO.Path]::GetFullPath($RoundTripCopyPath, $repoRoot)
	}

	& $powerPointScript @powerPointArguments
}
else {
	Write-Host "Step 4: Skipped PowerPoint validation." -ForegroundColor Yellow
}

if ($ExportSlides) {
	Write-Host "Step 5: Export slide images." -ForegroundColor Cyan
	$exportSlidesArguments = @{
		PptxPath = $resolvedOutputPath
		Format   = $SlideExportFormat
		CiSafe   = $CiSafe
	}

	if ($SlideExportDirectory) {
		$exportSlidesArguments.OutputDirectory = [System.IO.Path]::GetFullPath($SlideExportDirectory, $repoRoot)
	}

	& $exportSlidesScript @exportSlidesArguments
}

Write-Host "Smoke test passed for '$resolvedOutputPath'." -ForegroundColor Green

param(
    [string]$SamplesDirectory = "samples",
    [string]$OutputDirectory = "artifacts/smoke-tests",
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Debug",
    [switch]$IncludeRemoteSamples,
    [switch]$OnlyRemoteSamples,
    [switch]$IncludeCompatibilityGapSamples,
    [switch]$CiSafe,
    [switch]$SkipPowerPoint,
    [switch]$SkipRoundTripSave,
    [switch]$ContinueOnError
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-ThemeCssPath {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$MarkdownFile
    )

    $candidateNames = @(
        ("{0}.css" -f $MarkdownFile.BaseName),
        ($(if ($MarkdownFile.BaseName.EndsWith("-css")) { "{0}.css" -f $MarkdownFile.BaseName.Substring(0, $MarkdownFile.BaseName.Length - 4) } else { $null }))
    ) | Where-Object { $null -ne $_ } | Select-Object -Unique

    foreach ($candidateName in $candidateNames) {
        $candidatePath = Join-Path $MarkdownFile.DirectoryName $candidateName
        if (Test-Path $candidatePath) {
            return $candidatePath
        }
    }

    return $null
}

function Test-RequiresRemoteAssets {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$MarkdownFile
    )

    $content = [System.IO.File]::ReadAllText($MarkdownFile.FullName)
    return $content -match 'https?://'
}

function Test-IsCompatibilityGapSample {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$MarkdownFile
    )

    return $MarkdownFile.Name -eq "05-compatibility-gaps.md"
}

$repoRoot = Split-Path -Parent $PSScriptRoot
$smokeScript = Join-Path $PSScriptRoot "Invoke-PptxSmokeTest.ps1"
$resolvedSamplesDirectory = [System.IO.Path]::GetFullPath($SamplesDirectory, $repoRoot)
$resolvedOutputDirectory = [System.IO.Path]::GetFullPath($OutputDirectory, $repoRoot)
$includeRemoteSamplesEffective = $IncludeRemoteSamples -or $OnlyRemoteSamples

if (-not (Test-Path $resolvedSamplesDirectory)) {
    throw "Samples directory was not found: $resolvedSamplesDirectory"
}

New-Item -ItemType Directory -Path $resolvedOutputDirectory -Force | Out-Null

$sampleFiles = Get-ChildItem -Path $resolvedSamplesDirectory -File -Filter *.md |
Where-Object { $_.Name -ne "README.md" } |
Sort-Object Name

if (-not $sampleFiles) {
    throw "No sample Markdown files were found in '$resolvedSamplesDirectory'."
}

$failures = [System.Collections.Generic.List[string]]::new()
$executedCount = 0

foreach ($sampleFile in $sampleFiles) {
    $requiresRemoteAssets = Test-RequiresRemoteAssets -MarkdownFile $sampleFile

    if ($OnlyRemoteSamples -and -not $requiresRemoteAssets) {
        Write-Host "Skipping '$($sampleFile.Name)' because it does not require remote assets. Use the default mode to run the local smoke suite." -ForegroundColor Yellow
        continue
    }

    if ((Test-IsCompatibilityGapSample -MarkdownFile $sampleFile) -and -not $IncludeCompatibilityGapSamples) {
        Write-Host "Skipping '$($sampleFile.Name)' because it is a compatibility-gap repro sample. Use -IncludeCompatibilityGapSamples to include it." -ForegroundColor Yellow
        continue
    }

    if ($requiresRemoteAssets -and -not $includeRemoteSamplesEffective) {
        Write-Host "Skipping '$($sampleFile.Name)' because it requires remote assets. Use -IncludeRemoteSamples to enable remote smoke samples." -ForegroundColor Yellow
        continue
    }

    $outputPath = Join-Path $resolvedOutputDirectory ("{0}-{1}.pptx" -f $sampleFile.BaseName, $Configuration.ToLowerInvariant())
    $arguments = @{
        InputMarkdown = $sampleFile.FullName
        OutputPath    = $outputPath
        Configuration = $Configuration
    }

    $themeCssPath = Get-ThemeCssPath -MarkdownFile $sampleFile
    if ($themeCssPath) {
        $arguments.ThemeCss = $themeCssPath
    }

    if ($requiresRemoteAssets) {
        $arguments.AllowRemoteAssets = $true
    }

    if ($CiSafe) {
        $arguments.CiSafe = $true
    }

    if ($SkipPowerPoint) {
        $arguments.SkipPowerPoint = $true
    }

    if ($SkipRoundTripSave) {
        $arguments.SkipRoundTripSave = $true
    }

    try {
        $executedCount++
        Write-Host "Running smoke test for '$($sampleFile.Name)'..." -ForegroundColor Cyan
        & $smokeScript @arguments
    }
    catch {
        $failures.Add(("{0}: {1}" -f $sampleFile.Name, $_.Exception.Message))
        Write-Host "Smoke test failed for '$($sampleFile.Name)'." -ForegroundColor Red

        if (-not $ContinueOnError) {
            throw
        }
    }
}

if ($executedCount -eq 0) {
    throw "No smoke-test samples were selected from '$resolvedSamplesDirectory'."
}

if ($failures.Count -gt 0) {
    Write-Host "Smoke test failures:" -ForegroundColor Red
    foreach ($failure in $failures) {
        Write-Host (" - {0}" -f $failure) -ForegroundColor Red
    }

    throw ("{0} smoke test(s) failed." -f $failures.Count)
}

Write-Host ("Completed {0} smoke test(s) successfully." -f $executedCount) -ForegroundColor Green
param(
    [string]$SamplesDirectory = "samples",
    [string]$OutputDirectory = "artifacts/samples",
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Debug",
    [string]$Template,
    [switch]$Force
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

$repoRoot = Split-Path -Parent $PSScriptRoot
$generateScript = Join-Path $PSScriptRoot "Generate-LocalPptx.ps1"
$resolvedSamplesDirectory = [System.IO.Path]::GetFullPath($SamplesDirectory, $repoRoot)
$resolvedOutputDirectory = [System.IO.Path]::GetFullPath($OutputDirectory, $repoRoot)

if (-not (Test-Path $resolvedSamplesDirectory)) {
    throw "Samples directory was not found: $resolvedSamplesDirectory"
}

New-Item -ItemType Directory -Path $resolvedOutputDirectory -Force | Out-Null

$sampleFiles = Get-ChildItem -Path $resolvedSamplesDirectory -File -Filter *.md |
    Where-Object { $_.Name -ne "README.md" } |
    Sort-Object Name

if ($sampleFiles.Count -eq 0) {
    throw "No sample Markdown files were found in '$resolvedSamplesDirectory'."
}

foreach ($sampleFile in $sampleFiles) {
    $outputPath = Join-Path $resolvedOutputDirectory ("{0}.pptx" -f $sampleFile.BaseName)

    if ((-not $Force) -and (Test-Path $outputPath)) {
        Write-Host "Skipping '$($sampleFile.Name)' because '$outputPath' already exists. Use -Force to overwrite." -ForegroundColor Yellow
        continue
    }

    $arguments = @{
        InputMarkdown = $sampleFile.FullName
        OutputPath = $outputPath
        Configuration = $Configuration
    }

    $themeCssPath = Get-ThemeCssPath -MarkdownFile $sampleFile
    if ($themeCssPath) {
        $arguments.ThemeCss = $themeCssPath
        Write-Host "Using theme CSS '$themeCssPath' for '$($sampleFile.Name)'." -ForegroundColor Cyan
    }

    if ($Template) {
        $arguments.Template = [System.IO.Path]::GetFullPath($Template, $repoRoot)
    }

    Write-Host "Generating PPTX for '$($sampleFile.Name)'..." -ForegroundColor Cyan
    & $generateScript @arguments
}

Write-Host "Finished generating sample PPTX files in '$resolvedOutputDirectory'." -ForegroundColor Green

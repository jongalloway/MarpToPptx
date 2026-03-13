param(
    [Parameter(Mandatory = $true)]
    [string]$PptxPath,

    [string]$OutputDirectory,

    [ValidateSet("jpg", "png")]
    [string]$Format = "png",

    [switch]$CiSafe
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

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
$resolvedPptxPath = [System.IO.Path]::GetFullPath($PptxPath, $repoRoot)

if (-not (Test-Path $resolvedPptxPath -PathType Leaf)) {
    throw "PPTX file not found: $resolvedPptxPath"
}

if (-not $OutputDirectory) {
    $pptxBaseName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedPptxPath)
    $OutputDirectory = Join-Path (Join-Path (Join-Path $repoRoot "artifacts") "slide-exports") $pptxBaseName
}

$resolvedOutputDirectory = [System.IO.Path]::GetFullPath($OutputDirectory, $repoRoot)

$powerPointAvailable = Test-PowerPointAutomationAvailable

if (-not $powerPointAvailable) {
    if ($CiSafe) {
        Write-Host "CI-safe mode: PowerPoint COM automation is unavailable, so slide export will be skipped." -ForegroundColor Yellow
        return
    }

    if (-not $IsWindows) {
        throw "Slide export requires PowerPoint COM automation, which is only available on Windows with Microsoft PowerPoint installed."
    }

    throw "PowerPoint COM automation is not available. Ensure Microsoft PowerPoint is installed and accessible via COM interop."
}

New-Item -ItemType Directory -Path $resolvedOutputDirectory -Force | Out-Null

$filterName = $Format.ToUpperInvariant()
$ppt = $null
$presentation = $null

try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $presentation = $ppt.Presentations.Open($resolvedPptxPath, $false, $false, $false)
    $slideCount = $presentation.Slides.Count

    Write-Host ("Exporting {0} slide(s) from '{1}' as {2}..." -f $slideCount, $resolvedPptxPath, $filterName)

    for ($i = 1; $i -le $slideCount; $i++) {
        $paddedIndex = "{0:D3}" -f $i
        $slideFileName = "slide-{0}.{1}" -f $paddedIndex, $Format.ToLowerInvariant()
        $slidePath = Join-Path $resolvedOutputDirectory $slideFileName

        $presentation.Slides[$i].Export($slidePath, $filterName)
        Write-Host ("  Exported slide {0}/{1}: {2}" -f $i, $slideCount, $slideFileName)
    }

    Write-Host ("Exported {0} slide(s) to '{1}'." -f $slideCount, $resolvedOutputDirectory) -ForegroundColor Green
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}
finally {
    if ($presentation) {
        $presentation.Close() | Out-Null
    }

    if ($ppt) {
        $ppt.Quit()
    }
}

param(
    [string]$Template = "artifacts/real-world/vslive-lasvegas-ai/Next-Gen AI Apps with .NET.vsllv26-template.pptx",
    [string]$OutputDirectory = "artifacts/smoke-tests/e2e-update-run",
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",
    [switch]$CiSafe,
    [switch]$PassTemplateOnUpdate,
    [switch]$SkipPowerPointValidation
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Resolve-RepositoryFilePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot,

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    if ([System.IO.Path]::IsPathRooted($Path)) {
        $candidate = [System.IO.Path]::GetFullPath($Path)
    }
    else {
        $candidate = [System.IO.Path]::GetFullPath($Path, $RepositoryRoot)
    }

    if (-not (Test-Path $candidate -PathType Leaf)) {
        throw "$Description file was not found: $candidate"
    }

    return $candidate
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

function Write-DeckMarkdown {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Initial", "Updated")]
        [string]$Version
    )

    $content = if ($Version -eq "Initial") {
        @"
---
marp: true
layout: Title and Content
paginate: true
---

<!-- _layout: Title Slide -->
# Re-entrant Update Smoke Test

Author Name · March 2026

---

# Agenda

- Validate template-backed initial generation
- Preserve unmanaged PowerPoint edits
- Replace changed managed slide content

---

<!-- _layout: Section Header -->
# Deep Dive

Update workflow validation

---

# Details

Initial body text for the details slide.
"@
    }
    else {
        @"
---
marp: true
layout: Title and Content
paginate: true
---

<!-- slideId: re-entrant-update-smoke-test -->
<!-- _layout: Title Slide -->
# Re-entrant Update Smoke Test Updated

Author Name · March 2026

---

<!-- slideId: agenda -->
# Agenda

- Validate template-backed initial generation
- Preserve unmanaged PowerPoint slides
- Replace changed managed slide text
- Add a new managed slide during update

---

<!-- slideId: inserted-topic -->
<!-- _layout: Title and Content -->
# Inserted Topic

- Added after the first generation
- Should appear before the section header after update

---

<!-- slideId: deep-dive -->
<!-- _layout: Section Header -->
# Deep Dive

Update workflow validation

---

<!-- slideId: details -->
# Details

Updated body text for the details slide after the PowerPoint round-trip.
"@
    }

    Set-Content -Path $Path -Value $content -Encoding UTF8
}

function Invoke-LocalCli {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot,

        [Parameter(Mandatory = $true)]
        [string]$MarkdownPath,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath,

        [string]$TemplatePath,
        [string]$ExistingDeckPath
    )

    Push-Location $RepositoryRoot
    try {
        $arguments = @(
            "run"
            "--project"
            "src/MarpToPptx.Cli"
            "-c"
            $Configuration
            "--"
            $MarkdownPath
            "--write-slide-ids"
            "-o"
            $OutputPath
        )

        if ($TemplatePath) {
            $arguments += @("--template", $TemplatePath)
        }

        if ($ExistingDeckPath) {
            $arguments += @("--update-existing", $ExistingDeckPath)
        }

        Write-Host ("dotnet {0}" -f ($arguments -join " "))
        & dotnet @arguments

        if ($LASTEXITCODE -ne 0) {
            throw "CLI run failed with exit code $LASTEXITCODE"
        }
    }
    finally {
        Pop-Location
    }
}

function Invoke-OpenXmlValidation {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot,

        [Parameter(Mandatory = $true)]
        [string]$PptxPath
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
            throw "Open XML validation failed with exit code $LASTEXITCODE"
        }
    }
    finally {
        Pop-Location
    }
}

function Invoke-PowerPointManualEdit {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,

        [Parameter(Mandatory = $true)]
        [string]$EditedPath
    )

    Copy-Item $SourcePath $EditedPath -Force

    $ppt = $null
    $presentation = $null

    try {
        $ppt = New-Object -ComObject PowerPoint.Application
        $ppt.Visible = -1
        $presentation = $ppt.Presentations.Open($EditedPath, $false, $false, $false)

        $slide1 = $presentation.Slides.Item(1)
        $textShapes = @($slide1.Shapes | Where-Object { $_.HasTextFrame -eq -1 -and $_.TextFrame.HasText -eq -1 })
        if ($textShapes.Count -gt 0) {
            $textShapes[0].TextFrame.TextRange.Text = "Manual PowerPoint Edit Should Be Replaced"
        }

        $newSlide = $presentation.Slides.Add($presentation.Slides.Count + 1, 12)
        $textBox = $newSlide.Shapes.AddTextbox(1, 72, 72, 600, 80)
        $textBox.TextFrame.TextRange.Text = "User Added Slide"

        $presentation.Save()
    }
    finally {
        if ($presentation) {
            $presentation.Close() | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation) | Out-Null
        }

        if ($ppt) {
            $ppt.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Get-PowerPointSlideSummary {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PptxPath
    )

    $ppt = $null
    $presentation = $null

    try {
        $ppt = New-Object -ComObject PowerPoint.Application
        $ppt.Visible = -1
        $presentation = $ppt.Presentations.Open($PptxPath, $true, $false, $false)

        Write-Host ("SlideCount={0}" -f $presentation.Slides.Count)
        for ($index = 1; $index -le $presentation.Slides.Count; $index++) {
            $slide = $presentation.Slides.Item($index)
            $texts = @()
            foreach ($shape in @($slide.Shapes)) {
                if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
                    $texts += $shape.TextFrame.TextRange.Text.Replace("`r", " ").Replace("`n", " | ")
                }
            }

            Write-Host (("Slide {0}: " -f $index) + ($texts -join " || "))
        }
    }
    finally {
        if ($presentation) {
            $presentation.Close() | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation) | Out-Null
        }

        if ($ppt) {
            $ppt.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

$repoRoot = Split-Path -Parent $PSScriptRoot
$resolvedTemplate = Resolve-RepositoryFilePath -Path $Template -RepositoryRoot $repoRoot -Description "Template PPTX"
$resolvedOutputDirectory = [System.IO.Path]::GetFullPath($OutputDirectory, $repoRoot)
$powerPointScript = Join-Path $PSScriptRoot "Test-PowerPointOpen.ps1"

New-Item -ItemType Directory -Path $resolvedOutputDirectory -Force | Out-Null

$deckPath = Join-Path $resolvedOutputDirectory "deck.md"
$initialPath = Join-Path $resolvedOutputDirectory "initial.pptx"
$manualEditedPath = Join-Path $resolvedOutputDirectory "manual-edited.pptx"
$updatedPath = Join-Path $resolvedOutputDirectory "updated-from-existing.pptx"

$powerPointAvailable = Test-PowerPointAutomationAvailable
if (-not $powerPointAvailable -and -not $CiSafe) {
    throw "PowerPoint COM automation is required for this local E2E test. Re-run with -CiSafe only if you want to skip the PowerPoint steps."
}

if (-not $powerPointAvailable -and $CiSafe) {
    $SkipPowerPointValidation = $true
    Write-Host "CI-safe mode: PowerPoint COM automation is unavailable, so manual deck edits and PowerPoint-open validation will be skipped." -ForegroundColor Yellow
}

Write-Host "Step 1: Write the initial Marp deck." -ForegroundColor Cyan
Write-DeckMarkdown -Path $deckPath -Version Initial

Write-Host "Step 2: Generate the initial template-backed PPTX." -ForegroundColor Cyan
Invoke-LocalCli -RepositoryRoot $repoRoot -MarkdownPath $deckPath -OutputPath $initialPath -TemplatePath $resolvedTemplate
Invoke-OpenXmlValidation -RepositoryRoot $repoRoot -PptxPath $initialPath

if (-not $SkipPowerPointValidation) {
    Write-Host "Step 3: Simulate manual PowerPoint edits in the generated deck." -ForegroundColor Cyan
    Invoke-PowerPointManualEdit -SourcePath $initialPath -EditedPath $manualEditedPath
    & $powerPointScript -PptxPath $manualEditedPath
}
else {
    Copy-Item $initialPath $manualEditedPath -Force
    Write-Host "Step 3: Copied the generated deck without PowerPoint edits because PowerPoint validation is disabled." -ForegroundColor Yellow
}

Write-Host "Step 4: Update the Marp source." -ForegroundColor Cyan
Write-DeckMarkdown -Path $deckPath -Version Updated

Write-Host "Step 5: Regenerate against the existing edited deck." -ForegroundColor Cyan
$updateTemplate = if ($PassTemplateOnUpdate) { $resolvedTemplate } else { $null }
Invoke-LocalCli -RepositoryRoot $repoRoot -MarkdownPath $deckPath -OutputPath $updatedPath -ExistingDeckPath $manualEditedPath -TemplatePath $updateTemplate
Invoke-OpenXmlValidation -RepositoryRoot $repoRoot -PptxPath $updatedPath

if (-not $SkipPowerPointValidation) {
    Write-Host "Step 6: Verify the final output opens in PowerPoint." -ForegroundColor Cyan
    & $powerPointScript -PptxPath $updatedPath

    Write-Host "Step 7: Dump the final slide text summary." -ForegroundColor Cyan
    Get-PowerPointSlideSummary -PptxPath $updatedPath
}
else {
    Write-Host "Step 6: Skipped PowerPoint-open validation and slide summary." -ForegroundColor Yellow
}

Write-Host "E2E re-entrant update smoke test completed successfully." -ForegroundColor Green
Write-Host "Deck Markdown: $deckPath"
Write-Host "Initial deck:  $initialPath"
Write-Host "Edited deck:   $manualEditedPath"
Write-Host "Updated deck:  $updatedPath"
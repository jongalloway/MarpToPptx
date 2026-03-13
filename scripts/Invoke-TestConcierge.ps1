Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-RepositoryRoot {
    return (Split-Path -Parent $PSScriptRoot)
}

function Resolve-RepositoryPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot
    )

    return [System.IO.Path]::GetFullPath($Path, $RepositoryRoot)
}

function Get-DisplayPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot
    )

    $fullPath = Resolve-RepositoryPath -Path $Path -RepositoryRoot $RepositoryRoot
    return [System.IO.Path]::GetRelativePath($RepositoryRoot, $fullPath) -replace "\\", "/"
}

function Get-SampleMarkdownFiles {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot
    )

    $samplesDirectory = Join-Path $RepositoryRoot "samples"
    return Get-ChildItem -Path $samplesDirectory -File -Filter *.md |
        Where-Object { $_.Name -ne "README.md" } |
        Sort-Object Name
}

function Get-ThemeCssPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$MarkdownPath
    )

    $markdownFile = Get-Item $MarkdownPath
    $candidateNames = @(
        ("{0}.css" -f $markdownFile.BaseName),
        ($(if ($markdownFile.BaseName.EndsWith("-css")) { "{0}.css" -f $markdownFile.BaseName.Substring(0, $markdownFile.BaseName.Length - 4) } else { $null }))
    ) | Where-Object { $null -ne $_ } | Select-Object -Unique

    foreach ($candidateName in $candidateNames) {
        $candidatePath = Join-Path $markdownFile.DirectoryName $candidateName
        if (Test-Path $candidatePath -PathType Leaf) {
            return $candidatePath
        }
    }

    return $null
}

function Test-RequiresRemoteAssets {
    param(
        [Parameter(Mandatory = $true)]
        [string]$MarkdownPath
    )

    $content = [System.IO.File]::ReadAllText($MarkdownPath)
    return $content -match 'https?://'
}

function Read-MenuChoice {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Prompt,

        [Parameter(Mandatory = $true)]
        [object[]]$Options,

        [int]$DefaultIndex = 0
    )

    while ($true) {
        Write-Host ""
        Write-Host $Prompt -ForegroundColor Cyan
        for ($index = 0; $index -lt $Options.Count; $index++) {
            $marker = if ($index -eq $DefaultIndex) { "*" } else { " " }
            Write-Host (" {0} {1}. {2}" -f $marker, ($index + 1), $Options[$index].Label)
        }

        $selectionText = Read-Host ("Select 1-{0} (Enter for {1})" -f $Options.Count, ($DefaultIndex + 1))
        if ([string]::IsNullOrWhiteSpace($selectionText)) {
            return $Options[$DefaultIndex].Value
        }

        $selection = 0
        if ([int]::TryParse($selectionText, [ref]$selection) -and $selection -ge 1 -and $selection -le $Options.Count) {
            return $Options[$selection - 1].Value
        }

        Write-Warning "Enter a number from the list."
    }
}

function Read-YesNo {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Prompt,

        [bool]$Default = $true
    )

    $defaultLabel = if ($Default) { "Y/n" } else { "y/N" }

    while ($true) {
        $response = Read-Host ("{0} [{1}]" -f $Prompt, $defaultLabel)
        if ([string]::IsNullOrWhiteSpace($response)) {
            return $Default
        }

        switch ($response.Trim().ToLowerInvariant()) {
            "y" { return $true }
            "yes" { return $true }
            "n" { return $false }
            "no" { return $false }
            default { Write-Warning "Enter y or n." }
        }
    }
}

function Read-Text {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Prompt,

        [string]$Default = "",

        [switch]$AllowBlank
    )

    while ($true) {
        $responsePrompt = if ([string]::IsNullOrWhiteSpace($Default)) {
            $Prompt
        }
        else {
            "{0} [{1}]" -f $Prompt, $Default
        }

        $response = Read-Host $responsePrompt

        if ([string]::IsNullOrWhiteSpace($response)) {
            if ($AllowBlank) {
                return ""
            }

            if (-not [string]::IsNullOrWhiteSpace($Default)) {
                return $Default
            }

            Write-Warning "A value is required."
            continue
        }

        return $response.Trim()
    }
}

function Read-ExistingFilePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Prompt,

        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot,

        [string]$Default = "",

        [switch]$AllowBlank
    )

    while ($true) {
        $rawPath = Read-Text -Prompt $Prompt -Default $Default -AllowBlank:$AllowBlank
        if ([string]::IsNullOrWhiteSpace($rawPath)) {
            return $null
        }

        $resolvedPath = Resolve-RepositoryPath -Path $rawPath -RepositoryRoot $RepositoryRoot
        if (Test-Path $resolvedPath -PathType Leaf) {
            return $resolvedPath
        }

        Write-Warning ("File not found: {0}" -f $resolvedPath)
    }
}

function Read-ExistingPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Prompt,

        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot,

        [string]$Default = "",

        [switch]$AllowBlank
    )

    while ($true) {
        $rawPath = Read-Text -Prompt $Prompt -Default $Default -AllowBlank:$AllowBlank
        if ([string]::IsNullOrWhiteSpace($rawPath)) {
            return $null
        }

        $resolvedPath = Resolve-RepositoryPath -Path $rawPath -RepositoryRoot $RepositoryRoot
        if (Test-Path $resolvedPath) {
            return $resolvedPath
        }

        Write-Warning ("Path not found: {0}" -f $resolvedPath)
    }
}

function Read-MarkdownPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot
    )

    $sourceMode = Read-MenuChoice -Prompt "Choose the Markdown source." -Options @(
        [pscustomobject]@{ Label = "Select from samples/"; Value = "sample" },
        [pscustomobject]@{ Label = "Enter a custom Markdown path"; Value = "custom" }
    )

    if ($sourceMode -eq "sample") {
        $sampleFiles = Get-SampleMarkdownFiles -RepositoryRoot $RepositoryRoot
        $sampleOptions = foreach ($sampleFile in $sampleFiles) {
            [pscustomobject]@{
                Label = Get-DisplayPath -Path $sampleFile.FullName -RepositoryRoot $RepositoryRoot
                Value = $sampleFile.FullName
            }
        }

        return Read-MenuChoice -Prompt "Choose a sample deck." -Options $sampleOptions -DefaultIndex 0
    }

    return Read-ExistingFilePath -Prompt "Markdown path" -RepositoryRoot $RepositoryRoot
}

function Read-Configuration {
    return Read-MenuChoice -Prompt "Choose the .NET build configuration." -Options @(
        [pscustomobject]@{ Label = "Debug"; Value = "Debug" },
        [pscustomobject]@{ Label = "Release"; Value = "Release" }
    ) -DefaultIndex 0
}

function Read-ThemeCssSelection {
    param(
        [Parameter(Mandatory = $true)]
        [string]$MarkdownPath,

        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot
    )

    $detectedThemeCss = Get-ThemeCssPath -MarkdownPath $MarkdownPath
    if ($detectedThemeCss) {
        $detectedDisplayPath = Get-DisplayPath -Path $detectedThemeCss -RepositoryRoot $RepositoryRoot
        if (Read-YesNo -Prompt ("Use detected theme CSS '{0}'?" -f $detectedDisplayPath) -Default $true) {
            return $detectedThemeCss
        }
    }

    return Read-ExistingFilePath -Prompt "Theme CSS path (leave blank for none)" -RepositoryRoot $RepositoryRoot -AllowBlank
}

function Read-TemplateSelection {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot
    )

    return Read-ExistingFilePath -Prompt "Template PPTX path (leave blank for none)" -RepositoryRoot $RepositoryRoot -AllowBlank
}

function Get-DefaultGeneratedOutputPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$MarkdownPath,

        [Parameter(Mandatory = $true)]
        [string]$Configuration,

        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot
    )

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($MarkdownPath)
    return Join-Path (Join-Path (Join-Path $RepositoryRoot "artifacts") "samples") ("{0}-generated-{1}.pptx" -f $baseName, $Configuration.ToLowerInvariant())
}

function Read-OptionalOutputPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Prompt,

        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot,

        [string]$Default = "",

        [switch]$AllowBlank
    )

    $value = Read-Text -Prompt $Prompt -Default $Default -AllowBlank:$AllowBlank
    if ([string]::IsNullOrWhiteSpace($value)) {
        return ""
    }

    return Resolve-RepositoryPath -Path $value -RepositoryRoot $RepositoryRoot
}

function Read-CommonDeckOptions {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot
    )

    $markdownPath = Read-MarkdownPath -RepositoryRoot $RepositoryRoot
    $configuration = Read-Configuration
    $themeCss = Read-ThemeCssSelection -MarkdownPath $markdownPath -RepositoryRoot $RepositoryRoot
    $template = Read-TemplateSelection -RepositoryRoot $RepositoryRoot
    $requiresRemoteAssets = Test-RequiresRemoteAssets -MarkdownPath $markdownPath
    $allowRemoteAssets = if ($requiresRemoteAssets) {
        Read-YesNo -Prompt "Remote URLs were detected. Allow remote assets?" -Default $true
    }
    else {
        Read-YesNo -Prompt "Allow remote assets?" -Default $false
    }

    return [pscustomobject]@{
        MarkdownPath       = $markdownPath
        Configuration      = $configuration
        ThemeCss           = $themeCss
        Template           = $template
        AllowRemoteAssets  = $allowRemoteAssets
    }
}

function Invoke-SelectedScript {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ScriptPath,

        [Parameter(Mandatory = $true)]
        [hashtable]$Arguments,

        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot
    )

    Write-Host ""
    Write-Host ("Launching {0}..." -f (Get-DisplayPath -Path $ScriptPath -RepositoryRoot $RepositoryRoot)) -ForegroundColor Green
    & $ScriptPath @Arguments
}

function Start-GenerateSingle {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot
    )

    $options = Read-CommonDeckOptions -RepositoryRoot $RepositoryRoot
    $defaultOutputPath = Get-DefaultGeneratedOutputPath -MarkdownPath $options.MarkdownPath -Configuration $options.Configuration -RepositoryRoot $RepositoryRoot
    $outputPath = Read-OptionalOutputPath -Prompt "Output PPTX path" -RepositoryRoot $RepositoryRoot -Default $defaultOutputPath

    $arguments = @{
        InputMarkdown = $options.MarkdownPath
        OutputPath    = $outputPath
        Configuration = $options.Configuration
    }

    if ($options.ThemeCss) {
        $arguments.ThemeCss = $options.ThemeCss
    }

    if ($options.Template) {
        $arguments.Template = $options.Template
    }

    if ($options.AllowRemoteAssets) {
        $arguments.AllowRemoteAssets = $true
    }

    Invoke-SelectedScript -ScriptPath (Join-Path $PSScriptRoot "Generate-LocalPptx.ps1") -Arguments $arguments -RepositoryRoot $RepositoryRoot
}

function Read-SmokeMode {
    return Read-MenuChoice -Prompt "Choose the smoke-test behavior." -Options @(
        [pscustomobject]@{ Label = "Full smoke test (generate, validate, open in PowerPoint)"; Value = "full" },
        [pscustomobject]@{ Label = "CI-safe smoke test (open in PowerPoint only if automation is available)"; Value = "ci-safe" },
        [pscustomobject]@{ Label = "Validation only (skip the PowerPoint step)"; Value = "skip-powerpoint" }
    ) -DefaultIndex 1
}

function Add-CommonDeckArguments {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Arguments,

        [Parameter(Mandatory = $true)]
        [pscustomobject]$Options
    )

    $Arguments.InputMarkdown = $Options.MarkdownPath
    $Arguments.Configuration = $Options.Configuration

    if ($Options.ThemeCss) {
        $Arguments.ThemeCss = $Options.ThemeCss
    }

    if ($Options.Template) {
        $Arguments.Template = $Options.Template
    }

    if ($Options.AllowRemoteAssets) {
        $Arguments.AllowRemoteAssets = $true
    }
}

function Start-SmokeSingle {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot
    )

    $options = Read-CommonDeckOptions -RepositoryRoot $RepositoryRoot
    $smokeMode = Read-SmokeMode
    $outputPath = Read-OptionalOutputPath -Prompt "Output PPTX path (leave blank for the smoke-test default)" -RepositoryRoot $RepositoryRoot -AllowBlank
    $saveRoundTripCopy = if ($smokeMode -eq "skip-powerpoint") { $false } else { Read-YesNo -Prompt "Save a PowerPoint round-trip copy when PowerPoint runs?" -Default $true }

    $arguments = @{}
    Add-CommonDeckArguments -Arguments $arguments -Options $options

    if ($outputPath) {
        $arguments.OutputPath = $outputPath
    }

    switch ($smokeMode) {
        "ci-safe" { $arguments.CiSafe = $true }
        "skip-powerpoint" { $arguments.SkipPowerPoint = $true }
    }

    if (($smokeMode -ne "skip-powerpoint") -and -not $saveRoundTripCopy) {
        $arguments.SkipRoundTripSave = $true
    }

    Invoke-SelectedScript -ScriptPath (Join-Path $PSScriptRoot "Invoke-PptxSmokeTest.ps1") -Arguments $arguments -RepositoryRoot $RepositoryRoot
}

function Start-SmokeAll {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot
    )

    $configuration = Read-Configuration
    $suiteMode = Read-MenuChoice -Prompt "Choose which samples to run." -Options @(
        [pscustomobject]@{ Label = "Default local suite"; Value = "default" },
        [pscustomobject]@{ Label = "Include remote-asset samples"; Value = "include-remote" },
        [pscustomobject]@{ Label = "Only remote-asset samples"; Value = "only-remote" }
    ) -DefaultIndex 0
    $includeCompatibilityGapSamples = Read-YesNo -Prompt "Include compatibility-gap samples?" -Default $false
    $smokeMode = Read-SmokeMode
    $outputDirectory = Read-OptionalOutputPath -Prompt "Output directory" -RepositoryRoot $RepositoryRoot -Default (Join-Path (Join-Path $RepositoryRoot "artifacts") "smoke-tests")
    $saveRoundTripCopy = if ($smokeMode -eq "skip-powerpoint") { $false } else { Read-YesNo -Prompt "Save PowerPoint round-trip copies when PowerPoint runs?" -Default $true }
    $continueOnError = Read-YesNo -Prompt "Continue after failures?" -Default $true

    $arguments = @{
        Configuration = $configuration
        OutputDirectory = $outputDirectory
    }

    switch ($suiteMode) {
        "include-remote" { $arguments.IncludeRemoteSamples = $true }
        "only-remote" { $arguments.OnlyRemoteSamples = $true }
    }

    if ($includeCompatibilityGapSamples) {
        $arguments.IncludeCompatibilityGapSamples = $true
    }

    switch ($smokeMode) {
        "ci-safe" { $arguments.CiSafe = $true }
        "skip-powerpoint" { $arguments.SkipPowerPoint = $true }
    }

    if (($smokeMode -ne "skip-powerpoint") -and -not $saveRoundTripCopy) {
        $arguments.SkipRoundTripSave = $true
    }

    if ($continueOnError) {
        $arguments.ContinueOnError = $true
    }

    Invoke-SelectedScript -ScriptPath (Join-Path $PSScriptRoot "Invoke-AllPptxSmokeTests.ps1") -Arguments $arguments -RepositoryRoot $RepositoryRoot
}

function Start-ExpandPptx {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot
    )

    $pptxPath = Read-ExistingFilePath -Prompt "PPTX path to expand" -RepositoryRoot $RepositoryRoot
    $outputDirectory = Read-OptionalOutputPath -Prompt "Output directory (leave blank for the script default)" -RepositoryRoot $RepositoryRoot -AllowBlank
    $force = Read-YesNo -Prompt "Replace the output directory if it already exists?" -Default $true

    $arguments = @{
        PptxPath = $pptxPath
    }

    if ($outputDirectory) {
        $arguments.OutputDirectory = $outputDirectory
    }

    if ($force) {
        $arguments.Force = $true
    }

    Invoke-SelectedScript -ScriptPath (Join-Path $PSScriptRoot "Expand-Pptx.ps1") -Arguments $arguments -RepositoryRoot $RepositoryRoot
}

function Start-OpenInPowerPoint {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot
    )

    $pptxPath = Read-ExistingFilePath -Prompt "PPTX path to open in PowerPoint" -RepositoryRoot $RepositoryRoot
    $saveCopy = Read-YesNo -Prompt "Save a copy after PowerPoint opens it?" -Default $false

    $arguments = @{
        PptxPath = $pptxPath
    }

    if ($saveCopy) {
        $saveCopyAs = Read-OptionalOutputPath -Prompt "Save copy path" -RepositoryRoot $RepositoryRoot -Default (Join-Path (Split-Path -Parent $pptxPath) (([System.IO.Path]::GetFileNameWithoutExtension($pptxPath)) + "-powerpoint-resaved.pptx"))
        $arguments.SaveCopyAs = $saveCopyAs
    }

    Invoke-SelectedScript -ScriptPath (Join-Path $PSScriptRoot "Test-PowerPointOpen.ps1") -Arguments $arguments -RepositoryRoot $RepositoryRoot
}

function Start-ComparePptx {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot
    )

    $pathA = Read-ExistingPath -Prompt "First PPTX or expanded directory path" -RepositoryRoot $RepositoryRoot
    $pathB = Read-ExistingPath -Prompt "Second PPTX or expanded directory path" -RepositoryRoot $RepositoryRoot

    $arguments = @{
        PathA = $pathA
        PathB = $pathB
    }

    Invoke-SelectedScript -ScriptPath (Join-Path $PSScriptRoot "Compare-PptxStructure.ps1") -Arguments $arguments -RepositoryRoot $RepositoryRoot
}

function Start-ExportSlides {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot
    )

    $pptxPath = Read-ExistingFilePath -Prompt "PPTX path to export slides from" -RepositoryRoot $RepositoryRoot
    $format = Read-MenuChoice -Prompt "Choose the export image format." -Options @(
        [pscustomobject]@{ Label = "PNG"; Value = "png" },
        [pscustomobject]@{ Label = "JPG"; Value = "jpg" }
    ) -DefaultIndex 0
    $pptxBaseName = [System.IO.Path]::GetFileNameWithoutExtension($pptxPath)
    $defaultOutputDirectory = Join-Path (Join-Path (Join-Path $RepositoryRoot "artifacts") "slide-exports") $pptxBaseName
    $outputDirectory = Read-OptionalOutputPath -Prompt "Output directory" -RepositoryRoot $RepositoryRoot -Default $defaultOutputDirectory

    $arguments = @{
        PptxPath  = $pptxPath
        Format    = $format
    }

    if ($outputDirectory) {
        $arguments.OutputDirectory = $outputDirectory
    }

    Invoke-SelectedScript -ScriptPath (Join-Path $PSScriptRoot "Export-PptxSlides.ps1") -Arguments $arguments -RepositoryRoot $RepositoryRoot
}

$repositoryRoot = Get-RepositoryRoot

Write-Host "MarpToPptx test concierge" -ForegroundColor Magenta
Write-Host "Choose a workflow, answer a few questions, and the existing script will be launched for you." -ForegroundColor DarkGray

$action = Read-MenuChoice -Prompt "What do you want to do?" -Options @(
    [pscustomobject]@{ Label = "Generate one PPTX from Markdown"; Value = "generate-single" },
    [pscustomobject]@{ Label = "Run the smoke test for one deck"; Value = "smoke-single" },
    [pscustomobject]@{ Label = "Run the smoke test suite"; Value = "smoke-all" },
    [pscustomobject]@{ Label = "Export slide images from a PPTX"; Value = "export-slides" },
    [pscustomobject]@{ Label = "Expand a PPTX package into a directory"; Value = "expand" },
    [pscustomobject]@{ Label = "Open a PPTX in PowerPoint"; Value = "open" },
    [pscustomobject]@{ Label = "Compare two PPTX packages or directories"; Value = "compare" }
) -DefaultIndex 0

switch ($action) {
    "generate-single" { Start-GenerateSingle -RepositoryRoot $repositoryRoot }
    "smoke-single" { Start-SmokeSingle -RepositoryRoot $repositoryRoot }
    "smoke-all" { Start-SmokeAll -RepositoryRoot $repositoryRoot }
    "export-slides" { Start-ExportSlides -RepositoryRoot $repositoryRoot }
    "expand" { Start-ExpandPptx -RepositoryRoot $repositoryRoot }
    "open" { Start-OpenInPowerPoint -RepositoryRoot $repositoryRoot }
    "compare" { Start-ComparePptx -RepositoryRoot $repositoryRoot }
    default { throw "Unsupported action: $action" }
}
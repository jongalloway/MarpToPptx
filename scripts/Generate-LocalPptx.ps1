param(
    [Parameter(Mandatory = $true)]
    [string]$InputMarkdown,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [string]$ThemeCss,
    [string]$Template,
    [switch]$AllowRemoteAssets,
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Debug"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Resolve-RepositoryFilePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot,

        [string]$DefaultSubdirectory,

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    $candidates = New-Object System.Collections.Generic.List[string]

    if ([System.IO.Path]::IsPathRooted($Path)) {
        $candidates.Add([System.IO.Path]::GetFullPath($Path))
    }
    else {
        $candidates.Add([System.IO.Path]::GetFullPath($Path, $RepositoryRoot))

        if (-not [string]::IsNullOrWhiteSpace($DefaultSubdirectory)) {
            $candidates.Add([System.IO.Path]::GetFullPath((Join-Path $DefaultSubdirectory $Path), $RepositoryRoot))
        }
    }

    $uniqueCandidates = $candidates | Select-Object -Unique
    foreach ($candidate in $uniqueCandidates) {
        if (Test-Path $candidate -PathType Leaf) {
            return $candidate
        }
    }

    $tried = $uniqueCandidates -join [Environment]::NewLine
    throw "$Description file was not found. Tried:$([Environment]::NewLine)$tried"
}

$repoRoot = Split-Path -Parent $PSScriptRoot
$resolvedInputMarkdown = Resolve-RepositoryFilePath -Path $InputMarkdown -RepositoryRoot $repoRoot -DefaultSubdirectory "samples" -Description "Input Markdown"
$resolvedOutputPath = [System.IO.Path]::GetFullPath($OutputPath, $repoRoot)

if ($ThemeCss) {
    $resolvedThemeCss = Resolve-RepositoryFilePath -Path $ThemeCss -RepositoryRoot $repoRoot -DefaultSubdirectory "samples" -Description "Theme CSS"
}

if ($Template) {
    $resolvedTemplate = Resolve-RepositoryFilePath -Path $Template -RepositoryRoot $repoRoot -Description "Template PPTX"
}

Push-Location $repoRoot
try {
    $arguments = @(
        "run"
        "--project"
        "src/MarpToPptx.Cli"
        "-c"
        $Configuration
        "--"
        $resolvedInputMarkdown
        "-o"
        $resolvedOutputPath
    )

    if ($ThemeCss) {
        $arguments += @("--theme-css", $resolvedThemeCss)
    }

    if ($Template) {
        $arguments += @("--template", $resolvedTemplate)
    }

    if ($AllowRemoteAssets) {
        $arguments += "--allow-remote-assets"
    }

    Write-Host "Generating PPTX with local CLI project..."
    Write-Host ("dotnet {0}" -f ($arguments -join " "))
    & dotnet @arguments

    if ($LASTEXITCODE -ne 0) {
        throw "dotnet run failed with exit code $LASTEXITCODE"
    }

    Write-Host "Generated '$resolvedOutputPath'."
}
finally {
    Pop-Location
}

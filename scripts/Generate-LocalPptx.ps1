param(
    [Parameter(Mandatory = $true)]
    [string]$InputMarkdown,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [string]$ThemeCss,
    [string]$Template,
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Debug"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
Push-Location $repoRoot
try {
    $arguments = @(
        "run"
        "--project"
        "src/MarpToPptx.Cli"
        "-c"
        $Configuration
        "--"
        $InputMarkdown
        "-o"
        $OutputPath
    )

    if ($ThemeCss) {
        $arguments += @("--theme-css", $ThemeCss)
    }

    if ($Template) {
        $arguments += @("--template", $Template)
    }

    Write-Host "Generating PPTX with local CLI project..."
    Write-Host ("dotnet {0}" -f ($arguments -join " "))
    & dotnet @arguments

    if ($LASTEXITCODE -ne 0) {
        throw "dotnet run failed with exit code $LASTEXITCODE"
    }

    Write-Host "Generated '$OutputPath'."
}
finally {
    Pop-Location
}

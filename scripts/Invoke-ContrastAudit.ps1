param(
    [Parameter(Mandatory = $true)]
    [string]$PptxDirectory,

    [string]$ReportPath,

    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Debug",

    [switch]$NoBuild,

    [switch]$ContinueOnError
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$resolvedPptxDirectory = [System.IO.Path]::GetFullPath($PptxDirectory, $repoRoot)

if (-not (Test-Path $resolvedPptxDirectory)) {
    throw "PPTX directory was not found: $resolvedPptxDirectory"
}

$pptxFiles = Get-ChildItem -Path $resolvedPptxDirectory -Filter *.pptx | Sort-Object Name

if ($pptxFiles.Count -eq 0) {
    Write-Host "No PPTX files found in '$resolvedPptxDirectory'. Skipping contrast audit." -ForegroundColor Yellow
    return
}

Write-Host ("Running contrast audit on {0} PPTX file(s) in '{1}'..." -f $pptxFiles.Count, $resolvedPptxDirectory) -ForegroundColor Cyan

$auditResults = [System.Collections.Generic.List[PSCustomObject]]::new()
$failures = [System.Collections.Generic.List[string]]::new()
$errors = [System.Collections.Generic.List[string]]::new()

foreach ($pptxFile in $pptxFiles) {
    Write-Host "Auditing '$($pptxFile.Name)'..." -ForegroundColor Cyan

    $outputLines = [System.Collections.Generic.List[string]]::new()
    $exitCode = 0

    Push-Location $repoRoot
    try {
        $arguments = @(
            "run"
            "--project"
            "src/MarpToPptx.ContrastAuditor"
            "-c"
            $Configuration
        )

        if ($NoBuild) {
            $arguments += "--no-build"
        }

        $arguments += @("--", $pptxFile.FullName)

        Write-Host ("  dotnet {0}" -f ($arguments -join " "))

        & dotnet @arguments 2>&1 | ForEach-Object {
            $line = $_.ToString()
            Write-Host ("  {0}" -f $line)
            $outputLines.Add($line)
        }

        $exitCode = $LASTEXITCODE
    }
    finally {
        Pop-Location
    }

    $status = switch ($exitCode) {
        0 { "Pass" }
        2 { "Fail" }
        default { "Error" }
    }

    $auditResults.Add([PSCustomObject]@{
        File     = $pptxFile.Name
        Status   = $status
        ExitCode = $exitCode
        Output   = $outputLines -join "`n"
    })

    if ($exitCode -eq 0) {
        Write-Host ("  Contrast audit passed for '{0}'." -f $pptxFile.Name) -ForegroundColor Green
    }
    elseif ($exitCode -eq 2) {
        $failures.Add($pptxFile.Name)
        Write-Host ("  Contrast audit found failures in '{0}'." -f $pptxFile.Name) -ForegroundColor Red
        if (-not $ContinueOnError) {
            break
        }
    }
    else {
        $errors.Add($pptxFile.Name)
        Write-Host ("  Contrast audit reported an error for '{0}' (exit code {1})." -f $pptxFile.Name, $exitCode) -ForegroundColor Red
        if (-not $ContinueOnError) {
            break
        }
    }
}

if ($ReportPath) {
    $resolvedReportPath = [System.IO.Path]::GetFullPath($ReportPath, $repoRoot)
    $reportDirectory = Split-Path -Parent $resolvedReportPath
    if ($reportDirectory) {
        New-Item -ItemType Directory -Path $reportDirectory -Force | Out-Null
    }

    $auditResults | ConvertTo-Json -Depth 5 | Set-Content -Path $resolvedReportPath -Encoding UTF8
    Write-Host ("Contrast audit report written to '{0}'." -f $resolvedReportPath) -ForegroundColor Cyan
}

if ($errors.Count -gt 0) {
    Write-Host "Contrast audit errors:" -ForegroundColor Red
    foreach ($errorFile in $errors) {
        Write-Host (" - {0}" -f $errorFile) -ForegroundColor Red
    }
    throw ("{0} contrast audit(s) reported errors." -f $errors.Count)
}

if ($failures.Count -gt 0) {
    Write-Host "Contrast audit failures:" -ForegroundColor Red
    foreach ($failureFile in $failures) {
        Write-Host (" - {0}" -f $failureFile) -ForegroundColor Red
    }
    throw ("{0} contrast audit(s) found color-contrast failures." -f $failures.Count)
}

Write-Host ("Contrast audit passed for {0} PPTX file(s)." -f $pptxFiles.Count) -ForegroundColor Green

param(
    [Parameter(Mandatory = $true)]
    [string]$PptxPath,

    [string]$OutputDirectory,

    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$resolvedPptxPath = Resolve-Path $PptxPath
if (-not $OutputDirectory) {
    $OutputDirectory = [System.IO.Path]::Combine(
        [System.IO.Path]::GetDirectoryName($resolvedPptxPath),
        [System.IO.Path]::GetFileNameWithoutExtension($resolvedPptxPath) + "_unzipped")
}

if ((Test-Path $OutputDirectory) -and -not $Force) {
    throw "Output directory '$OutputDirectory' already exists. Use -Force to replace it."
}

if (Test-Path $OutputDirectory) {
    Remove-Item $OutputDirectory -Recurse -Force
}

$tempZipPath = Join-Path ([System.IO.Path]::GetTempPath()) (([System.IO.Path]::GetFileNameWithoutExtension($resolvedPptxPath)) + "-" + [guid]::NewGuid().ToString("N") + ".zip")
Copy-Item $resolvedPptxPath $tempZipPath -Force

try {
    Expand-Archive -Path $tempZipPath -DestinationPath $OutputDirectory -Force
    Write-Host "Expanded '$resolvedPptxPath' to '$OutputDirectory'."
}
finally {
    if (Test-Path $tempZipPath) {
        Remove-Item $tempZipPath -Force
    }
}

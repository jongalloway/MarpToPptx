param(
    [Parameter(Mandatory = $true)]
    [string]$PathA,

    [Parameter(Mandatory = $true)]
    [string]$PathB
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Expand-IfNeeded {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPath
    )

    $resolved = (Resolve-Path $InputPath).Path
    if (Test-Path $resolved -PathType Container) {
        return @{
            Directory = $resolved
            Temporary = $false
        }
    }

    if ([System.IO.Path]::GetExtension($resolved) -ne ".pptx") {
        throw "Expected a .pptx file or an expanded directory: $resolved"
    }

    $tempDirectory = Join-Path ([System.IO.Path]::GetTempPath()) ("pptx-compare-" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $tempDirectory -Force | Out-Null

    $zipPath = Join-Path ([System.IO.Path]::GetTempPath()) (([System.IO.Path]::GetFileNameWithoutExtension($resolved)) + "-" + [guid]::NewGuid().ToString("N") + ".zip")
    Copy-Item $resolved $zipPath -Force
    try {
        Expand-Archive -Path $zipPath -DestinationPath $tempDirectory -Force
    }
    finally {
        Remove-Item $zipPath -Force
    }

    return @{
        Directory = $tempDirectory
        Temporary = $true
    }
}

function Get-RelativeFiles {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Root
    )

    Get-ChildItem -Path $Root -Recurse -File |
        ForEach-Object {
            $_.FullName.Substring($Root.Length).TrimStart("\\", "/") -replace "\\", "/"
        } |
        Sort-Object -Unique
}

$packageA = Expand-IfNeeded -InputPath $PathA
$packageB = Expand-IfNeeded -InputPath $PathB

try {
    $filesA = Get-RelativeFiles -Root $packageA.Directory
    $filesB = Get-RelativeFiles -Root $packageB.Directory

    Write-Host "Only in A:"
    Compare-Object -ReferenceObject $filesA -DifferenceObject $filesB |
        Where-Object SideIndicator -eq "<=" |
        ForEach-Object { Write-Host "  $($_.InputObject)" }

    Write-Host "Only in B:"
    Compare-Object -ReferenceObject $filesA -DifferenceObject $filesB |
        Where-Object SideIndicator -eq "=>" |
        ForEach-Object { Write-Host "  $($_.InputObject)" }

    $interestingPaths = @(
        "[Content_Types].xml",
        "_rels/.rels",
        "ppt/_rels/presentation.xml.rels",
        "ppt/slideMasters/_rels/slideMaster1.xml.rels"
    )

    $interestingPaths += $filesA + $filesB |
        Where-Object { $_ -like "ppt/slides/_rels/*.rels" -or $_ -like "ppt/slideLayouts/_rels/*.rels" } |
        Sort-Object -Unique

    Write-Host "Content differences in key XML files:"
    foreach ($relativePath in ($interestingPaths | Sort-Object -Unique)) {
        $fileA = Join-Path $packageA.Directory $relativePath
        $fileB = Join-Path $packageB.Directory $relativePath
        if ((Test-Path $fileA) -and (Test-Path $fileB)) {
            $contentA = [System.IO.File]::ReadAllText($fileA)
            $contentB = [System.IO.File]::ReadAllText($fileB)
            if (-not [string]::Equals($contentA, $contentB, [System.StringComparison]::Ordinal)) {
                Write-Host "  differs: $relativePath"
            }
        }
    }
}
finally {
    if ($packageA.Temporary -and (Test-Path $packageA.Directory)) {
        Remove-Item $packageA.Directory -Recurse -Force
    }

    if ($packageB.Temporary -and (Test-Path $packageB.Directory)) {
        Remove-Item $packageB.Directory -Recurse -Force
    }
}

param(
    [Parameter(Mandatory = $true)]
    [string]$PptxPath,

    [string]$SaveCopyAs
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$resolvedPptxPath = (Resolve-Path $PptxPath).Path
$ppt = $null
$presentation = $null

try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $presentation = $ppt.Presentations.Open($resolvedPptxPath, $false, $false, $false)

    Write-Host "Opened '$resolvedPptxPath' successfully in PowerPoint."

    if ($SaveCopyAs) {
        $saveCopyPath = [System.IO.Path]::GetFullPath($SaveCopyAs)
        $saveDirectory = Split-Path -Parent $SaveCopyAs
        if ($saveDirectory) {
            New-Item -ItemType Directory -Path $saveDirectory -Force | Out-Null
        }

        $presentation.SaveCopyAs($saveCopyPath)
        Write-Host "Saved copy to '$saveCopyPath'."
    }
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

param(
	[Parameter(Mandatory = $true)]
	[string]$InputMarkdown,

	[string]$OutputPath,
	[string]$ThemeCss,
	[string]$Template,
	[ValidateSet("Debug", "Release")]
	[string]$Configuration = "Debug",
	[switch]$CiSafe,
	[switch]$SkipPowerPoint,
	[switch]$SkipRoundTripSave,
	[string]$RoundTripCopyPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$script:OpenXmlAssemblies = @()

function Get-OpenXmlType {
	param(
		[Parameter(Mandatory = $true)]
		[string]$TypeName
	)

	foreach ($assembly in $script:OpenXmlAssemblies) {
		$type = $assembly.GetType($TypeName, $false)
		if ($null -ne $type) {
			return $type
		}
	}

	throw "Unable to find type [$TypeName] in the loaded Open XML assemblies."
}

function Get-PackageAssemblyPath {
	param(
		[Parameter(Mandatory = $true)]
		[hashtable]$AssetsData,

		[Parameter(Mandatory = $true)]
		[string]$PackageId,

		[Parameter(Mandatory = $true)]
		[string]$FileName,

		[Parameter(Mandatory = $true)]
		[string]$NuGetPackagesRoot
	)

	$libraryEntry = $AssetsData["libraries"].GetEnumerator() |
		Where-Object { $_.Key -like "$PackageId/*" } |
		Select-Object -First 1

	if ($null -eq $libraryEntry) {
		throw "Unable to find package '$PackageId' in project.assets.json."
	}

	$packageFiles = @($libraryEntry.Value["files"])
	$preferredFrameworks = @()
	$runtimeVersionMajor = [System.Environment]::Version.Major

	if ($runtimeVersionMajor -ge 10) {
		$preferredFrameworks += "net10.0"
	}

	if ($runtimeVersionMajor -ge 8) {
		$preferredFrameworks += "net8.0"
	}

	if ($runtimeVersionMajor -ge 6) {
		$preferredFrameworks += "net6.0"
	}

	$preferredFrameworks += "netstandard2.0"

	$relativeAssemblyPath = $null
	foreach ($framework in $preferredFrameworks | Select-Object -Unique) {
		$candidatePath = "lib/$framework/$FileName"
		if ($packageFiles -contains $candidatePath) {
			$relativeAssemblyPath = $candidatePath
			break
		}
	}

	if ($null -eq $relativeAssemblyPath) {
		throw "Unable to find a PowerShell-compatible runtime assembly '$FileName' for package '$PackageId'."
	}

	return Join-Path (Join-Path $NuGetPackagesRoot $libraryEntry.Value["path"]) $relativeAssemblyPath
}

function Import-OpenXmlValidationAssemblies {
	param(
		[Parameter(Mandatory = $true)]
		[string]$CliOutputDirectory
	)

	$cliProjectDirectory = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $CliOutputDirectory))
	$assetsFilePath = Join-Path (Join-Path $cliProjectDirectory "obj") "project.assets.json"
	if (-not (Test-Path $assetsFilePath)) {
		throw "Unable to locate project.assets.json at '$assetsFilePath'."
	}

	$nuGetPackagesRoot = if ($env:NUGET_PACKAGES) {
		$env:NUGET_PACKAGES
	}
	elseif ($HOME) {
		Join-Path $HOME ".nuget/packages"
	}
	else {
		throw "Unable to determine the NuGet global-packages directory."
	}

	$assetsData = Get-Content -Path $assetsFilePath -Raw | ConvertFrom-Json -AsHashtable
	$assemblyPaths = @(
		(Get-PackageAssemblyPath -AssetsData $assetsData -PackageId "System.IO.Packaging" -FileName "System.IO.Packaging.dll" -NuGetPackagesRoot $nuGetPackagesRoot),
		(Get-PackageAssemblyPath -AssetsData $assetsData -PackageId "DocumentFormat.OpenXml.Framework" -FileName "DocumentFormat.OpenXml.Framework.dll" -NuGetPackagesRoot $nuGetPackagesRoot),
		(Get-PackageAssemblyPath -AssetsData $assetsData -PackageId "DocumentFormat.OpenXml" -FileName "DocumentFormat.OpenXml.dll" -NuGetPackagesRoot $nuGetPackagesRoot)
	)

	$loadedAssemblies = @()

	foreach ($assemblyPath in $assemblyPaths) {
		if (-not (Test-Path $assemblyPath)) {
			throw "Required validation assembly was not found: $assemblyPath"
		}

		$assemblyName = [System.Reflection.AssemblyName]::GetAssemblyName($assemblyPath)
		$assembly = [AppDomain]::CurrentDomain.GetAssemblies() |
			Where-Object { $_.FullName -eq $assemblyName.FullName } |
			Select-Object -First 1
		if ($null -eq $assembly) {
			$tempDirectory = Join-Path ([System.IO.Path]::GetTempPath()) ("MarpToPptx-openxml-{0}" -f ([System.Guid]::NewGuid().ToString("N")))
			New-Item -ItemType Directory -Path $tempDirectory -Force | Out-Null
			$tempAssemblyPath = Join-Path $tempDirectory ([System.IO.Path]::GetFileName($assemblyPath))
			Copy-Item -Path $assemblyPath -Destination $tempAssemblyPath -Force
			$assembly = [System.Reflection.Assembly]::LoadFrom($tempAssemblyPath)
		}

		$loadedAssemblies += $assembly
	}

	$script:OpenXmlAssemblies = $loadedAssemblies
}

function Test-OpenXmlPackage {
	param(
		[Parameter(Mandatory = $true)]
		[string]$PptxPath
	)

	$validationErrors = @()
	$document = $null
	$presentationDocumentType = Get-OpenXmlType -TypeName "DocumentFormat.OpenXml.Packaging.PresentationDocument"
	$openXmlValidatorType = Get-OpenXmlType -TypeName "DocumentFormat.OpenXml.Validation.OpenXmlValidator"
	$openMethod = $presentationDocumentType.GetMethods() |
		Where-Object {
			$parameters = $_.GetParameters()
			$_.Name -eq "Open" -and
			$parameters.Count -eq 2 -and
			$parameters[0].ParameterType -eq [string] -and
			$parameters[1].ParameterType -eq [bool]
		} |
		Select-Object -First 1

	if ($null -eq $openMethod) {
		throw "Unable to find PresentationDocument.Open(string, bool)."
	}

	try {
		$document = $openMethod.Invoke($null, @($PptxPath, $false))
		$validator = [System.Activator]::CreateInstance($openXmlValidatorType)
		$validationErrors = @($validator.Validate($document))
	}
	finally {
		if ($document) {
			$document.Dispose()
		}
	}

	if ($validationErrors.Count -gt 0) {
		Write-Host "Open XML validation failed:" -ForegroundColor Red
		foreach ($validationError in $validationErrors) {
			$pathText = if ($validationError.Path) {
				if ($validationError.Path | Get-Member -Name XPath -ErrorAction SilentlyContinue) {
					$validationError.Path.XPath
				}
				else {
					$validationError.Path.ToString()
				}
			}
			else {
				"<unknown path>"
			}

			Write-Host ("- {0} at {1}" -f $validationError.Description, $pathText)
		}

		throw "Open XML validation found $($validationErrors.Count) error(s)."
	}

	Write-Host "Open XML validation passed." -ForegroundColor Green
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

$repoRoot = Split-Path -Parent $PSScriptRoot
$generateScript = Join-Path $PSScriptRoot "Generate-LocalPptx.ps1"
$powerPointScript = Join-Path $PSScriptRoot "Test-PowerPointOpen.ps1"

if (-not $OutputPath) {
	$inputBaseName = [System.IO.Path]::GetFileNameWithoutExtension($InputMarkdown)
	$smokeTestDirectory = Join-Path (Join-Path $repoRoot "artifacts") "smoke-tests"
	$OutputPath = Join-Path $smokeTestDirectory ("{0}-{1}.pptx" -f $inputBaseName, $Configuration.ToLowerInvariant())
}

$resolvedOutputPath = [System.IO.Path]::GetFullPath($OutputPath, $repoRoot)
$outputDirectory = Split-Path -Parent $resolvedOutputPath
if ($outputDirectory) {
	New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
}

if (-not $SkipPowerPoint -and -not $SkipRoundTripSave -and -not $RoundTripCopyPath) {
	$outputBaseName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedOutputPath)
	$RoundTripCopyPath = Join-Path $outputDirectory ("{0}-roundtrip.pptx" -f $outputBaseName)
}

$powerPointAvailable = Test-PowerPointAutomationAvailable

if ($CiSafe -and -not $SkipPowerPoint -and -not $powerPointAvailable) {
	$SkipPowerPoint = $true
	Write-Host "CI-safe mode: PowerPoint COM automation is unavailable, so the smoke test will stop after Open XML validation." -ForegroundColor Yellow
}

Write-Host "Step 1: Generate PPTX from the local workspace code." -ForegroundColor Cyan
$generateArguments = @{
	InputMarkdown = $InputMarkdown
	OutputPath = $resolvedOutputPath
	Configuration = $Configuration
}

if ($ThemeCss) {
	$generateArguments.ThemeCss = $ThemeCss
}

if ($Template) {
	$generateArguments.Template = $Template
}

& $generateScript @generateArguments

$cliOutputDirectory = Join-Path (Join-Path (Join-Path (Join-Path $repoRoot "src") "MarpToPptx.Cli") "bin") (Join-Path $Configuration "net10.0")
if (-not (Test-Path $cliOutputDirectory)) {
	throw "CLI output directory was not found after generation: $cliOutputDirectory"
}

Write-Host "Step 2: Validate the generated package with Open XML SDK." -ForegroundColor Cyan
Import-OpenXmlValidationAssemblies -CliOutputDirectory $cliOutputDirectory
Test-OpenXmlPackage -PptxPath $resolvedOutputPath

if (-not $SkipPowerPoint) {
	Write-Host "Step 3: Open the package in PowerPoint." -ForegroundColor Cyan
	$powerPointArguments = @{
		PptxPath = $resolvedOutputPath
	}

	if (-not $SkipRoundTripSave -and $RoundTripCopyPath) {
		$powerPointArguments.SaveCopyAs = [System.IO.Path]::GetFullPath($RoundTripCopyPath, $repoRoot)
	}

	& $powerPointScript @powerPointArguments
}
else {
	Write-Host "Step 3: Skipped PowerPoint validation." -ForegroundColor Yellow
}

Write-Host "Smoke test passed for '$resolvedOutputPath'." -ForegroundColor Green

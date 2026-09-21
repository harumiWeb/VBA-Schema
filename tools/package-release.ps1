[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Tag,

    [string]$OutputDirectory = "artifacts/release"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$tagPattern = '^v(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)$'
if ($Tag -notmatch $tagPattern) {
    throw "Release tag must match vMAJOR.MINOR.PATCH without leading zeroes: $Tag"
}

$repoRoot = [System.IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot))
if ([System.IO.Path]::IsPathRooted($OutputDirectory)) {
    $outputDirectoryPath = [System.IO.Path]::GetFullPath($OutputDirectory)
} else {
    $outputDirectoryPath = [System.IO.Path]::GetFullPath((Join-Path $repoRoot $OutputDirectory))
}

$repoPrefix = $repoRoot.TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar
if (-not $outputDirectoryPath.StartsWith($repoPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Release package output must be inside the repository: $outputDirectoryPath"
}
if ($outputDirectoryPath.Equals($repoRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Release package output must not be the repository root"
}

$sourceNames = @("Schema.bas", "VSchema.cls", "VValidationResult.cls")
$stagingPath = Join-Path $repoRoot "dist\VBA-Release"
$stageScript = Join-Path $PSScriptRoot "release-stage.ps1"
$verifyScript = Join-Path $PSScriptRoot "release-verify.ps1"

function Get-Sha256 {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $algorithm = [System.Security.Cryptography.SHA256]::Create()
    try {
        return (($algorithm.ComputeHash($bytes) | ForEach-Object { $_.ToString("x2") }) -join "")
    } finally {
        $algorithm.Dispose()
    }
}

& $stageScript -Destination $stagingPath
& $verifyScript -Destination $stagingPath

New-Item -ItemType Directory -Path $outputDirectoryPath -Force | Out-Null
$packagePath = Join-Path $outputDirectoryPath "VBA-Release-$Tag.zip"
$checksumPath = "$packagePath.sha256"
if (Test-Path -LiteralPath $packagePath) {
    Remove-Item -LiteralPath $packagePath -Force
}
if (Test-Path -LiteralPath $checksumPath) {
    Remove-Item -LiteralPath $checksumPath -Force
}

$temporaryRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("vba-release-" + [guid]::NewGuid().ToString("N"))
$packageRoot = Join-Path $temporaryRoot "VBA-Release"
try {
    New-Item -ItemType Directory -Path $packageRoot -Force | Out-Null
    foreach ($sourceName in $sourceNames) {
        $sourcePath = Join-Path $stagingPath $sourceName
        if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
            throw "Staged release source is missing: $sourceName"
        }
        Copy-Item -LiteralPath $sourcePath -Destination (Join-Path $packageRoot $sourceName)
    }

    Compress-Archive -LiteralPath $packageRoot -DestinationPath $packagePath -CompressionLevel Optimal -Force
    if (-not (Test-Path -LiteralPath $packagePath -PathType Leaf)) {
        throw "Release package was not created: $packagePath"
    }

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::OpenRead($packagePath)
    try {
        $actualEntries = @(
            $archive.Entries |
                Where-Object { -not $_.FullName.EndsWith("/") } |
                ForEach-Object { $_.FullName.Replace("\", "/") }
        )
        $expectedEntries = @(
            "VBA-Release/Schema.bas",
            "VBA-Release/VSchema.cls",
            "VBA-Release/VValidationResult.cls"
        )
        $actualSorted = @($actualEntries | Sort-Object)
        $expectedSorted = @($expectedEntries | Sort-Object)
        if (($actualSorted -join "`n") -ne ($expectedSorted -join "`n")) {
            throw "Release package entries are not the expected three modules: $($actualSorted -join ', ')"
        }
    } finally {
        $archive.Dispose()
    }
} finally {
    if (Test-Path -LiteralPath $temporaryRoot) {
        Remove-Item -LiteralPath $temporaryRoot -Recurse -Force
    }
}

$hash = Get-Sha256 $packagePath
[System.IO.File]::WriteAllText(
    $checksumPath,
    "$hash  $([System.IO.Path]::GetFileName($packagePath))$([Environment]::NewLine)",
    [System.Text.Encoding]::ASCII
)
Write-Output "Packaged VBA-Release: $packagePath"
Write-Output "SHA-256: $hash"
Write-Output "SHA-256 checksum: $checksumPath"

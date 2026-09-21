[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = [System.IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot))
$releaseScript = Join-Path $PSScriptRoot "release-stage.ps1"
$protectedDestinations = @(
    ".",
    "src\modules",
    "src\classes",
    ".git",
    ".xlflow",
    "build"
)
$sourceFiles = @(
    "src\modules\Schema.bas",
    "src\classes\VSchema.cls",
    "src\classes\VValidationResult.cls"
)

function Get-Sha256 {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $algorithm = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.IO.File]::ReadAllBytes($Path)
        return (($algorithm.ComputeHash($bytes) | ForEach-Object { $_.ToString("x2") }) -join "")
    } finally {
        $algorithm.Dispose()
    }
}

$beforeHashes = @{}
foreach ($relativePath in $sourceFiles) {
    $path = Join-Path $repoRoot $relativePath
    $beforeHashes[$relativePath] = Get-Sha256 $path
}

foreach ($relativeDestination in $protectedDestinations) {
    $destination = [System.IO.Path]::GetFullPath((Join-Path $repoRoot $relativeDestination))
    $previousErrorAction = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        $output = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $releaseScript -Destination $destination 2>&1 | Out-String
        $exitCode = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $previousErrorAction
    }
    if ($exitCode -eq 0) {
        throw "release-stage unexpectedly accepted protected destination: $relativeDestination"
    }
    if ($output -notmatch "must not overlap|inside the repository") {
        throw "release-stage rejected protected destination for an unexpected reason: $relativeDestination`n$output"
    }
}

foreach ($relativePath in $sourceFiles) {
    $path = Join-Path $repoRoot $relativePath
    $afterHash = Get-Sha256 $path
    if ($afterHash -ne $beforeHashes[$relativePath]) {
        throw "release-stage safety probe changed a production source file: $relativePath"
    }
}

Write-Output "Release staging safety passed: protected destinations were rejected without mutating production sources."

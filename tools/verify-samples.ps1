[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = [System.IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot))
$sampleRoot = Join-Path $repoRoot "sample"
$samplePaths = @(
    Get-ChildItem -LiteralPath $sampleRoot -Recurse -File -Filter "*.bas" |
        Sort-Object FullName |
        ForEach-Object { $_.FullName }
)
if ($samplePaths.Count -eq 0) {
    throw "No sample VBA source was found under $sampleRoot"
}

$configuredXlflowPath = [Environment]::GetEnvironmentVariable("VBA_SCHEMA_XLFLOW_BIN")
if ([string]::IsNullOrWhiteSpace($configuredXlflowPath)) {
    $configuredXlflowPath = (Get-Command xlflow -CommandType Application -ErrorAction Stop).Source
}
if (-not (Test-Path -LiteralPath $configuredXlflowPath -PathType Leaf)) {
    throw "xlflow executable not found: $configuredXlflowPath"
}

function Invoke-Xlflow {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Arguments
    )

    & $configuredXlflowPath @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "xlflow command failed with exit code $($LASTEXITCODE): $($Arguments -join ' ')"
    }
}

$temporaryRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("vba-schema-samples-" + [guid]::NewGuid().ToString("N"))
try {
    Write-Output "Checking sample format: $($samplePaths.Count) file(s)"
    Invoke-Xlflow (@("fmt", "--check") + $samplePaths)

    New-Item -ItemType Directory -Path (Join-Path $temporaryRoot "src\modules") -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $temporaryRoot "src\classes") -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $repoRoot "xlflow.toml") -Destination $temporaryRoot
    Copy-Item -LiteralPath (Join-Path $repoRoot "src\modules\Schema.bas") -Destination (Join-Path $temporaryRoot "src\modules")
    Copy-Item -LiteralPath (Join-Path $repoRoot "src\classes\VSchema.cls") -Destination (Join-Path $temporaryRoot "src\classes")
    Copy-Item -LiteralPath (Join-Path $repoRoot "src\classes\VValidationResult.cls") -Destination (Join-Path $temporaryRoot "src\classes")
    foreach ($samplePath in $samplePaths) {
        Copy-Item -LiteralPath $samplePath -Destination (Join-Path $temporaryRoot "src\modules")
    }

    Push-Location $temporaryRoot
    try {
        Write-Output "Checking sample lint and analysis in an isolated source project"
        Invoke-Xlflow @("lint", "--json")
        Invoke-Xlflow @("analyze", "--json")
    } finally {
        Pop-Location
    }
} finally {
    if (Test-Path -LiteralPath $temporaryRoot) {
        Remove-Item -LiteralPath $temporaryRoot -Recurse -Force
    }
}

Write-Output "Sample source verification passed. Excel/VBE compile is a separate Windows gate."

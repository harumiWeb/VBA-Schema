[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$targetPaths = @(
    "src\modules\Schema.bas",
    "src\classes\VSchema.cls",
    "src\classes\VValidationResult.cls",
    "src\modules\Tests\PublicApiCompile.bas"
)

foreach ($testFile in @("src\modules\Tests")) {
    $testDirectory = Join-Path $repoRoot $testFile
    if (Test-Path -LiteralPath $testDirectory -PathType Container) {
        Get-ChildItem -LiteralPath $testDirectory -File -Filter "Test*.bas" |
            ForEach-Object { $targetPaths += Join-Path $testFile $_.Name }
    }
}

$benchmarkDirectory = Join-Path $repoRoot "src\modules\Benchmarks"
if (Test-Path -LiteralPath $benchmarkDirectory -PathType Container) {
    Get-ChildItem -LiteralPath $benchmarkDirectory -File -Filter "*.bas" |
        ForEach-Object { $targetPaths += Join-Path "src\modules\Benchmarks" $_.Name }
}

$existingPaths = @($targetPaths | Where-Object { Test-Path -LiteralPath (Join-Path $repoRoot $_) -PathType Leaf })
if ($existingPaths.Count -eq 0) {
    Write-Output "No VBA-Schema production or focused test source exists yet; format check deferred."
    exit 0
}

$configuredXlflowPath = [Environment]::GetEnvironmentVariable("VBA_SCHEMA_XLFLOW_BIN")
if ([string]::IsNullOrWhiteSpace($configuredXlflowPath)) {
    $configuredXlflowPath = (Get-Command xlflow -CommandType Application -ErrorAction Stop).Source
}
if (-not (Test-Path -LiteralPath $configuredXlflowPath -PathType Leaf)) {
    throw "xlflow executable not found: $configuredXlflowPath"
}

Write-Output "Using xlflow: $configuredXlflowPath"
Push-Location $repoRoot
try {
    & $configuredXlflowPath fmt --check $existingPaths
    exit $LASTEXITCODE
} finally {
    Pop-Location
}

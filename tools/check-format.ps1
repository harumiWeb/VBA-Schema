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

foreach ($testFile in @("src\modules\Tests", "src\modules\Benchmarks")) {
    $testDirectory = Join-Path $repoRoot $testFile
    if (Test-Path -LiteralPath $testDirectory -PathType Container) {
        Get-ChildItem -LiteralPath $testDirectory -File -Filter "Test*.bas" |
            ForEach-Object { $targetPaths += Join-Path $testFile $_.Name }
    }
}

$existingPaths = @($targetPaths | Where-Object { Test-Path -LiteralPath (Join-Path $repoRoot $_) -PathType Leaf })
if ($existingPaths.Count -eq 0) {
    Write-Output "No VBA-Schema production or focused test source exists yet; format check deferred."
    exit 0
}

$xlflowCommand = Get-Command xlflow -CommandType Application -ErrorAction Stop
Push-Location $repoRoot
try {
    & $xlflowCommand.Source fmt --check $existingPaths
    exit $LASTEXITCODE
} finally {
    Pop-Location
}

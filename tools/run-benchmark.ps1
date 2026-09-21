[CmdletBinding()]
param(
    [string]$OutputDirectory = "artifacts\benchmarks",
    [string]$BaselinePath = "benchmarks\windows-x64-baseline.json",
    [switch]$UpdateBaseline
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "benchmark-environment.ps1")

$repoRoot = [System.IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot))
$outputDirectoryPath = if ([System.IO.Path]::IsPathRooted($OutputDirectory)) {
    [System.IO.Path]::GetFullPath($OutputDirectory)
} else {
    [System.IO.Path]::GetFullPath((Join-Path $repoRoot $OutputDirectory))
}
$baselinePath = if ([System.IO.Path]::IsPathRooted($BaselinePath)) {
    [System.IO.Path]::GetFullPath($BaselinePath)
} else {
    [System.IO.Path]::GetFullPath((Join-Path $repoRoot $BaselinePath))
}

$xlflowCommand = Get-Command xlflow -CommandType Application -ErrorAction Stop
$sessionStarted = $false

function Invoke-Xlflow {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Arguments
    )

    $previousErrorAction = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        $output = & $xlflowCommand.Source @Arguments 2>&1 | Out-String -Width 65535
        $exitCode = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $previousErrorAction
    }
    if ($exitCode -ne 0) {
        throw "xlflow failed (exit $exitCode): $($Arguments -join ' ')`n$output"
    }
    return $output
}

function Convert-XlflowJson {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Output,
        [Parameter(Mandatory = $true)]
        [string]$CommandName
    )

    $jsonStart = $Output.IndexOf("{", [System.StringComparison]::Ordinal)
    if ($jsonStart -lt 0) {
        throw "xlflow $CommandName did not return JSON.`n$Output"
    }
    try {
        return $Output.Substring($jsonStart).Trim() | ConvertFrom-Json
    } catch {
        throw "xlflow $CommandName returned invalid JSON: $($_.Exception.Message)`n$Output"
    }
}

function Save-Utf8NoBomJson {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [object]$Value
    )

    $parent = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
    $json = $Value | ConvertTo-Json -Depth 8
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $json + [Environment]::NewLine, $utf8NoBom)
}

function Get-FixtureMap {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Benchmark
    )

    $map = @{}
    foreach ($fixture in @($Benchmark.fixtures)) {
        $map[[string]$fixture.fixture_id] = $fixture
    }
    return $map
}

function Assert-BenchmarkContract {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Benchmark
    )

    if ([int]$Benchmark.schema_version -ne 1) {
        throw "Unexpected benchmark schema_version: $($Benchmark.schema_version)"
    }
    if ([int]$Benchmark.warmup_iterations -ne 1 -or [int]$Benchmark.iterations -ne 5) {
        throw "Benchmark iteration contract changed unexpectedly."
    }

    $expected = @{
        object_1000_fields_success = @{ validation_count = 50000; issue_count = 0; target_ms = 500 }
        object_1000_fields_issues = @{ validation_count = 50000; issue_count = 5000; target_ms = 0 }
        scalar_10000_success = @{ validation_count = 50000; issue_count = 0; target_ms = 1000 }
        scalar_10000_issues = @{ validation_count = 50000; issue_count = 50000; target_ms = 0 }
    }
    $fixtures = Get-FixtureMap $Benchmark
    if ($fixtures.Count -ne $expected.Count) {
        throw "Expected $($expected.Count) benchmark fixtures, got $($fixtures.Count)."
    }

    foreach ($fixtureId in $expected.Keys) {
        if (-not $fixtures.ContainsKey($fixtureId)) {
            throw "Missing benchmark fixture: $fixtureId"
        }
        $fixture = $fixtures[$fixtureId]
        $contract = $expected[$fixtureId]
        if ([int]$fixture.warmup_iterations -ne 1 -or [int]$fixture.iterations -ne 5) {
            throw "Unexpected iteration metadata for $fixtureId."
        }
        if ([int]$fixture.raw_ms.Count -ne 5) {
            throw "Expected five raw samples for $fixtureId."
        }
        if ([int]$fixture.validation_count -ne $contract.validation_count) {
            throw "Unexpected validation_count for $fixtureId."
        }
        if ([int]$fixture.issue_count -ne $contract.issue_count) {
            throw "Unexpected issue_count for $fixtureId."
        }
        if ([double]$fixture.target_ms -ne [double]$contract.target_ms) {
            throw "Unexpected target_ms for $fixtureId."
        }
        if ([double]$fixture.median_ms -lt 0) {
            throw "Negative median_ms for $fixtureId."
        }
    }
    return $fixtures
}

function Compare-BenchmarkBaseline {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Benchmark,
        [Parameter(Mandatory = $true)]
        [hashtable]$CurrentFixtures,
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [switch]$AllowContractChange
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return [ordered]@{ status = "missing"; path = $Path; regressions = @() }
    }

    $baseline = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
    if ([int]$baseline.schema_version -ne 1) {
        throw "Unsupported benchmark baseline schema_version: $($baseline.schema_version)"
    }
    $baselineFixtures = Get-FixtureMap $baseline
    $regressions = @()
    foreach ($fixtureId in $CurrentFixtures.Keys) {
        if (-not $baselineFixtures.ContainsKey($fixtureId)) {
            if ($AllowContractChange) {
                return [ordered]@{ status = "skipped_for_update"; path = $Path; regressions = @() }
            }
            throw "Baseline is missing fixture: $fixtureId"
        }
        $current = $CurrentFixtures[$fixtureId]
        $previous = $baselineFixtures[$fixtureId]
        if ([int]$current.validation_count -ne [int]$previous.validation_count -or [int]$current.issue_count -ne [int]$previous.issue_count) {
            if ($AllowContractChange) {
                return [ordered]@{ status = "skipped_for_update"; path = $Path; regressions = @() }
            }
            throw "Benchmark fixture contract differs from baseline: $fixtureId"
        }
        $previousMedian = [double]$previous.median_ms
        if ($previousMedian -gt 0 -and [double]$current.median_ms -gt ($previousMedian * 1.25)) {
            $regressions += [ordered]@{
                fixture_id = $fixtureId
                baseline_median_ms = $previousMedian
                current_median_ms = [double]$current.median_ms
                allowed_median_ms = $previousMedian * 1.25
            }
        }
    }
    return [ordered]@{
        status = if ($regressions.Count -eq 0) { "pass" } else { "failed" }
        path = $Path
        regressions = $regressions
    }
}

try {
    $statusOutput = Invoke-Xlflow @("status", "--json")
    $status = Convert-XlflowJson $statusOutput "status"
    if ($status.coordination.recovery_required) {
        throw "xlflow reports recovery_required; resolve recovery before running benchmark."
    }
    if ($status.coordination.busy -or $status.session.active -or $status.session.workbook_open) {
        throw "A workbook or xlflow session is already active. Benchmark requires exclusive managed-session ownership."
    }

    Invoke-Xlflow @("lint", "--json") | Out-Null
    Invoke-Xlflow @("analyze", "--json") | Out-Null
    Invoke-Xlflow @("session", "start", "--json") | Out-Null
    $sessionStarted = $true
    # A newly started managed session may contain an older workbook while the
    # push state cache says the source is unchanged. Omit --fast so the source
    # is imported unconditionally into this benchmark-owned session.
    Invoke-Xlflow @("push", "--session", "--no-save", "--json") | Out-Null
    $runOutput = Invoke-Xlflow @("run", "ValidationBenchmarks.RunValidationBenchmark", "--diagnostic", "--headless", "--session", "--no-save", "--json")

    $match = [regex]::Match($runOutput, "VBA_SCHEMA_BENCHMARK_JSON=(\{.*\})")
    if (-not $match.Success) {
        throw "Benchmark macro did not emit VBA_SCHEMA_BENCHMARK_JSON.`n$runOutput"
    }
    $benchmark = $match.Groups[1].Value | ConvertFrom-Json
    $fixtures = Assert-BenchmarkContract $benchmark

    $absoluteFailures = @()
    foreach ($fixture in @($benchmark.fixtures)) {
        if ([double]$fixture.target_ms -gt 0 -and [double]$fixture.median_ms -ge [double]$fixture.target_ms) {
            $absoluteFailures += [string]$fixture.fixture_id
        }
    }

    $environmentDecision = Get-BenchmarkEnvironmentDecision ([string]$benchmark.office_bitness)
    if ($environmentDecision.status -eq "pass") {
        $baselineComparison = Compare-BenchmarkBaseline $benchmark $fixtures $baselinePath -AllowContractChange:$UpdateBaseline
    } else {
        $baselineComparison = [ordered]@{
            status = "not_run"
            path = $baselinePath
            regressions = @()
        }
    }
    if ($UpdateBaseline -and $baselineComparison.status -eq "failed") {
        # An explicit baseline refresh is the reviewer's escape hatch after a
        # stable fixture/metric-set decision. Absolute targets still gate it.
        $baselineComparison = [ordered]@{
            status = "skipped_for_update"
            path = $baselineComparison.path
            regressions = $baselineComparison.regressions
        }
    }
    $absoluteTargetStatus = "pass"
    if ($absoluteFailures.Count -gt 0) {
        $absoluteTargetStatus = "failed"
    }

    $timestamp = Get-Date -Format "yyyyMMdd-HHmmssfff"
    $outputPath = Join-Path $outputDirectoryPath "$timestamp-$($environmentDecision.report_suffix).json"
    $report = [ordered]@{
        schema_version = 1
        generated_at = $benchmark.generated_at
        machine = $benchmark.machine
        office_bitness = $benchmark.office_bitness
        excel_version = $benchmark.excel_version
        excel_build = $benchmark.excel_build
        command = "task benchmark"
        warmup_iterations = $benchmark.warmup_iterations
        iterations = $benchmark.iterations
        fixtures = $benchmark.fixtures
        verification = [ordered]@{
            environment = $environmentDecision.status
            environment_error = $environmentDecision.error
            absolute_targets = $absoluteTargetStatus
            baseline = $baselineComparison
        }
    }
    Save-Utf8NoBomJson $outputPath $report

    if ($environmentDecision.status -ne "pass") {
        throw "$($environmentDecision.error) Report: $outputPath"
    }
    if ($absoluteFailures.Count -gt 0) {
        throw "Absolute performance target failed: $($absoluteFailures -join ', '). Report: $outputPath"
    }
    if ($baselineComparison.status -eq "failed") {
        throw "Benchmark regression exceeded 25 percent: $($baselineComparison.regressions.fixture_id -join ', '). Report: $outputPath"
    }

    if ($UpdateBaseline) {
        $baseline = [ordered]@{
            schema_version = 1
            environment_class = "windows-x64-office"
            fixture_contract = [ordered]@{
                warmup_iterations = $benchmark.warmup_iterations
                iterations = $benchmark.iterations
            }
            fixtures = @($benchmark.fixtures | ForEach-Object {
                [ordered]@{
                    fixture_id = $_.fixture_id
                    median_ms = $_.median_ms
                    validation_count = $_.validation_count
                    issue_count = $_.issue_count
                    warmup_iterations = $_.warmup_iterations
                    iterations = $_.iterations
                }
            })
        }
        Save-Utf8NoBomJson $baselinePath $baseline
        Write-Output "Updated benchmark baseline: $baselinePath"
    }

    Write-Output "Benchmark passed: $outputPath"
    Write-Output "Baseline comparison: $($baselineComparison.status)"
} finally {
    if ($sessionStarted) {
        try {
            Invoke-Xlflow @("session", "stop", "--discard", "--json") | Out-Null
        } catch {
            Write-Error "Failed to stop benchmark-owned xlflow session: $($_.Exception.Message)"
            throw
        }
    }
}

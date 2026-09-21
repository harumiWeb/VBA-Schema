[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "benchmark-environment.ps1")

$x64 = Get-BenchmarkEnvironmentDecision "x64"
if ($x64.status -ne "pass" -or $x64.report_suffix -ne "windows-x64" -or $null -ne $x64.error) {
    throw "x64 benchmark environment was not accepted as the measured environment."
}

$x86 = Get-BenchmarkEnvironmentDecision "x86"
if ($x86.status -ne "failed" -or $x86.report_suffix -ne "unsupported-x86" -or $x86.error -notmatch "requires Windows 64-bit Excel") {
    throw "x86 benchmark environment was not rejected with an actionable error."
}

$unknown = Get-BenchmarkEnvironmentDecision ""
if ($unknown.status -ne "failed" -or $unknown.report_suffix -ne "unsupported-unknown") {
    throw "Unknown benchmark environment was not rejected safely."
}

Write-Output "Benchmark environment guard passed: only x64 is eligible for the Windows x64 baseline."

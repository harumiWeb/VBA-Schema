Set-StrictMode -Version Latest

function Get-BenchmarkEnvironmentDecision {
    [CmdletBinding()]
    param(
        [AllowEmptyString()]
        [string]$OfficeBitness
    )

    $normalizedBitness = if ([string]::IsNullOrWhiteSpace($OfficeBitness)) {
        "unknown"
    } else {
        $OfficeBitness.Trim()
    }

    if ($normalizedBitness -ieq "x64") {
        return [pscustomobject]@{
            status        = "pass"
            report_suffix = "windows-x64"
            error         = $null
        }
    }

    $safeBitness = $normalizedBitness -replace "[^A-Za-z0-9_-]", "_"
    return [pscustomobject]@{
        status        = "failed"
        report_suffix = "unsupported-$safeBitness"
        error         = "Benchmark requires Windows 64-bit Excel; detected office_bitness='$normalizedBitness'."
    }
}

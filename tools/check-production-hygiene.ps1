[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = [System.IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot))
$productionPaths = @(
    "src\modules\Schema.bas",
    "src\classes\VSchema.cls",
    "src\classes\VValidationResult.cls"
)

$rules = @(
    @{ Name = "Windows API declaration"; Pattern = '^\s*(?:Public\s+|Private\s+|Friend\s+)?Declare(?:\s+PtrSafe)?\b' },
    @{ Name = "pointer-size dependent type"; Pattern = '\b(?:LongPtr|LongLong)\b' },
    @{ Name = "unqualified Excel/host reference"; Pattern = '\b(?:Application|Excel)\s*\.' },
    @{ Name = "UI state mutation"; Pattern = '\.(?:Select|Activate)\b' },
    @{ Name = "broad error suppression"; Pattern = '^\s*On\s+Error\s+Resume\s+Next\b' }
)

$violations = New-Object System.Collections.Generic.List[string]
foreach ($relativePath in $productionPaths) {
    $path = Join-Path $repoRoot $relativePath
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Production source is missing: $path"
    }

    $lineNumber = 0
    foreach ($line in Get-Content -LiteralPath $path) {
        $lineNumber++
        $trimmed = $line.TrimStart()
        if ($trimmed.StartsWith("'")) {
            continue
        }
        foreach ($rule in $rules) {
            if ($line -match $rule.Pattern) {
                $violations.Add("$relativePath`:$lineNumber [$($rule.Name)] $($line.Trim())")
            }
        }
    }
}

if ($violations.Count -gt 0) {
    throw "Production hygiene audit failed:`n$($violations -join "`n")"
}

Write-Output "Production hygiene audit passed: $($productionPaths.Count) core files checked."

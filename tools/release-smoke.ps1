[CmdletBinding()]
param(
    [string]$Payload = "dist/VBA-Schema"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = [System.IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot))
$payloadPath = if ([System.IO.Path]::IsPathRooted($Payload)) {
    [System.IO.Path]::GetFullPath($Payload)
} else {
    [System.IO.Path]::GetFullPath((Join-Path $repoRoot $Payload))
}

$payloadNames = @("Schema.bas", "VSchema.cls", "VValidationResult.cls")
foreach ($name in $payloadNames) {
    $path = Join-Path $payloadPath $name
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Release payload is missing: $path. Run release-stage first."
    }
}

$xlflowCommand = Get-Command xlflow -CommandType Application -ErrorAction Stop
$probeName = "vba-schema-release-smoke-" + [guid]::NewGuid().ToString("N")
$probeRoot = Join-Path ([System.IO.Path]::GetTempPath()) $probeName
$probeWorkbook = Join-Path $probeRoot "build\Book.xlsm"
$locationDepth = 0

function Invoke-Xlflow {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Arguments
    )

    $previousErrorAction = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        $output = & $xlflowCommand.Source @Arguments 2>&1 | Out-String
        $exitCode = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $previousErrorAction
    }
    if ($exitCode -ne 0) {
        throw "xlflow failed (exit $exitCode): $($Arguments -join ' ')`n$output"
    }
    return $output
}

function Assert-WithinProbe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $resolved = [System.IO.Path]::GetFullPath($Path)
    $prefix = $probeRoot.TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar
    if (-not $resolved.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to mutate path outside release smoke probe: $resolved"
    }
    return $resolved
}

try {
    New-Item -ItemType Directory -Path $probeRoot -Force | Out-Null
    Push-Location $probeRoot
    $locationDepth = 1

    Invoke-Xlflow @("new", ".\Book.xlsm", "--no-update-check", "--json") | Out-Null
    if (-not (Test-Path -LiteralPath $probeWorkbook -PathType Leaf)) {
        throw "xlflow new did not create the probe workbook: $probeWorkbook"
    }

    $modulesPath = Assert-WithinProbe (Join-Path $probeRoot "src\modules")
    $classesPath = Assert-WithinProbe (Join-Path $probeRoot "src\classes")
    if (Test-Path -LiteralPath $modulesPath) {
        Remove-Item -LiteralPath $modulesPath -Recurse -Force
    }
    if (Test-Path -LiteralPath $classesPath) {
        Remove-Item -LiteralPath $classesPath -Recurse -Force
    }
    New-Item -ItemType Directory -Path $modulesPath -Force | Out-Null
    New-Item -ItemType Directory -Path $classesPath -Force | Out-Null

    foreach ($name in $payloadNames) {
        $destination = if ($name.EndsWith(".cls")) {
            Join-Path $classesPath $name
        } else {
            Join-Path $modulesPath $name
        }
        Copy-Item -LiteralPath (Join-Path $payloadPath $name) -Destination $destination -Force
    }

    # Do not use --fast here: a fresh probe can share xlflow's push-state cache
    # with another checkout and incorrectly skip the first import.
    $initialPush = Invoke-Xlflow @("push", "--json")
    if ($initialPush -notmatch 'imported 3 source file\(s\)') {
        throw "Release payload push did not import exactly three source files.`n$initialPush"
    }

    $smokeSource = @'
Attribute VB_Name = "ReleaseSmoke"
Option Explicit

Public Sub ScalarSmoke()
    Dim successResult As VValidationResult
    Set successResult = Schema.Text().Min(2).SafeParse("ok")
    If Not successResult.Success Then
        Err.Raise vbObjectError + 2301, "ReleaseSmoke.ScalarSmoke", "expected scalar success"
    End If

    Dim failureResult As VValidationResult
    Set failureResult = Schema.Text().Min(3).SafeParse("x")
    If failureResult.Success Then
        Err.Raise vbObjectError + 2302, "ReleaseSmoke.ScalarSmoke", "expected scalar failure"
    End If
    If failureResult.Issues.Count <> 1 Then
        Err.Raise vbObjectError + 2303, "ReleaseSmoke.ScalarSmoke", "expected one issue"
    End If
    If Len(failureResult.ErrorText) = 0 Then
        Err.Raise vbObjectError + 2304, "ReleaseSmoke.ScalarSmoke", "expected ErrorText"
    End If
End Sub
'@
    $smokePath = Assert-WithinProbe (Join-Path $modulesPath "ReleaseSmoke.bas")
    Set-Content -LiteralPath $smokePath -Value $smokeSource -Encoding ascii
    Invoke-Xlflow @("push", "--json") | Out-Null
    Invoke-Xlflow @("run", "ReleaseSmoke.ScalarSmoke", "--diagnostic", "--headless", "--json") | Out-Null

    Remove-Item -LiteralPath $smokePath -Force
    $finalPush = Invoke-Xlflow @("push", "--json")
    if ($finalPush -notmatch 'imported 3 source file\(s\)') {
        throw "Final release payload push did not return to exactly three source files.`n$finalPush"
    }

    Invoke-Xlflow @("pull", "--json") | Out-Null
    $moduleFiles = @(Get-ChildItem -LiteralPath $modulesPath -File -Recurse | Where-Object { $_.Extension -in @(".bas", ".cls") })
    $classFiles = @(Get-ChildItem -LiteralPath $classesPath -File -Recurse | Where-Object { $_.Extension -in @(".bas", ".cls") })
    $allFiles = @($moduleFiles) + @($classFiles)
    $componentNames = @($allFiles | ForEach-Object { $_.Name } | Sort-Object)
    $expectedNames = @($payloadNames | Sort-Object)
    if (($componentNames -join "|") -ne ($expectedNames -join "|")) {
        throw "Non-document component set differs from the three-file release payload. Actual: $($componentNames -join ', '); expected: $($expectedNames -join ', ')"
    }

    Write-Output "Release smoke passed: imported, compiled, executed scalar smoke, and verified 3 non-document components. Probe: $probeRoot"
} finally {
    while ($locationDepth -gt 0) {
        Pop-Location
        $locationDepth--
    }
    if (Test-Path -LiteralPath $probeRoot) {
        Remove-Item -LiteralPath $probeRoot -Recurse -Force
    }
}

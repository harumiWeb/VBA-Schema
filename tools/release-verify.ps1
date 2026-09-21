[CmdletBinding()]
param(
    [string]$Destination = "dist/VBA-Schema"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = [System.IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot))
if ([System.IO.Path]::IsPathRooted($Destination)) {
    $destinationPath = [System.IO.Path]::GetFullPath($Destination)
} else {
    $destinationPath = [System.IO.Path]::GetFullPath((Join-Path $repoRoot $Destination))
}

$expectedNames = @("Schema.bas", "VSchema.cls", "VValidationResult.cls")
if (-not (Test-Path -LiteralPath $destinationPath -PathType Container)) {
    throw "Release payload directory does not exist: $destinationPath"
}

$actualFiles = @(Get-ChildItem -LiteralPath $destinationPath -File -Recurse | ForEach-Object { $_.FullName.Substring($destinationPath.Length + 1) })
$unexpected = @($actualFiles | Where-Object { $_ -notin $expectedNames })
$missing = @($expectedNames | Where-Object { $_ -notin $actualFiles })
if ($unexpected.Count -gt 0) {
    throw "Unexpected release payload files: $($unexpected -join ', ')"
}
if ($missing.Count -gt 0) {
    throw "Missing release payload files: $($missing -join ', ')"
}

foreach ($name in $expectedNames) {
    $path = Join-Path $destinationPath $name
    $bytes = [System.IO.File]::ReadAllBytes($path)
    if ($bytes -contains [byte]0x0D) {
        throw "Release payload must use LF line endings: $name"
    }
    if ($bytes.Length -ge 3 -and $bytes[0] -eq [byte]0xEF -and $bytes[1] -eq [byte]0xBB -and $bytes[2] -eq [byte]0xBF) {
        throw "Release payload must be UTF-8 without BOM: $name"
    }
    $utf8Strict = New-Object -TypeName System.Text.UTF8Encoding -ArgumentList $false, $true
    try {
        $text = $utf8Strict.GetString($bytes)
    } catch {
        throw "Release payload is not valid UTF-8: $name"
    }
    if ($text -notmatch '(?m)^Attribute VB_Name = "[^"]+"') {
        throw "Missing Attribute VB_Name in release payload: $name"
    }
    if ($name.EndsWith(".cls") -and $text -notmatch '(?m)^VERSION 1\.0 CLASS') {
        throw "Missing class header in release payload: $name"
    }
}

Write-Output "Verified 3-file import payload: $destinationPath"

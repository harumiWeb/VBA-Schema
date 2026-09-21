[CmdletBinding()]
param(
    [string]$Destination = "build\Book.xlsm"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
if ([System.IO.Path]::IsPathRooted($Destination)) {
    $destinationPath = [System.IO.Path]::GetFullPath($Destination)
} else {
    $destinationPath = [System.IO.Path]::GetFullPath((Join-Path $repoRoot $Destination))
}

if (Test-Path -LiteralPath $destinationPath -PathType Leaf) {
    throw "Refusing to overwrite the tracked or user-owned development workbook: $destinationPath"
}

$xlflowCommand = Get-Command xlflow -CommandType Application -ErrorAction Stop
$probeName = "vba-schema-bootstrap-" + [guid]::NewGuid().ToString("N")
$probeRoot = Join-Path ([System.IO.Path]::GetTempPath()) $probeName
$probeWorkbook = Join-Path $probeRoot "build\Book.xlsm"

try {
    New-Item -ItemType Directory -Path $probeRoot -Force | Out-Null
    Push-Location $probeRoot
    try {
        & $xlflowCommand.Source new ".\Book.xlsm" --no-update-check --json
        if ($LASTEXITCODE -ne 0) {
            throw "xlflow new failed with exit code $LASTEXITCODE"
        }
        if (-not (Test-Path -LiteralPath $probeWorkbook -PathType Leaf)) {
            throw "xlflow new did not create the expected workbook: $probeWorkbook"
        }
    } finally {
        Pop-Location
    }

    $destinationDirectory = Split-Path -Parent $destinationPath
    New-Item -ItemType Directory -Path $destinationDirectory -Force | Out-Null
    Copy-Item -LiteralPath $probeWorkbook -Destination $destinationPath
    Write-Output "Provisioned workbook: $destinationPath"
} finally {
    if (Test-Path -LiteralPath $probeRoot) {
        Remove-Item -LiteralPath $probeRoot -Recurse -Force
    }
}

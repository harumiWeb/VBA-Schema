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

$repoPrefix = $repoRoot.TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar
if (-not $destinationPath.StartsWith($repoPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Release staging destination must be inside the repository: $destinationPath"
}
if ($destinationPath.Equals($repoRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Release staging destination must not overlap the repository root: $destinationPath"
}

$destinationPrefix = $destinationPath.TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar
$protectedRoots = @(
    (Join-Path $repoRoot "src"),
    (Join-Path $repoRoot ".git"),
    (Join-Path $repoRoot ".xlflow"),
    (Join-Path $repoRoot "build")
)
foreach ($protectedRootCandidate in $protectedRoots) {
    $protectedRoot = [System.IO.Path]::GetFullPath($protectedRootCandidate)
    $protectedPrefix = $protectedRoot.TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar
    $destinationOverlapsProtectedRoot =
        $destinationPath.Equals($protectedRoot, [System.StringComparison]::OrdinalIgnoreCase) -or
        $destinationPath.StartsWith($protectedPrefix, [System.StringComparison]::OrdinalIgnoreCase) -or
        $protectedRoot.StartsWith($destinationPrefix, [System.StringComparison]::OrdinalIgnoreCase)
    if ($destinationOverlapsProtectedRoot) {
        throw "Release staging destination must not overlap protected repository path: $destinationPath"
    }
}

$sourceFiles = @(
    "src\modules\Schema.bas",
    "src\classes\VSchema.cls",
    "src\classes\VValidationResult.cls"
)

foreach ($relativePath in $sourceFiles) {
    $sourcePath = Join-Path $repoRoot $relativePath
    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
        throw "Release source is missing: $relativePath"
    }
}

if (Test-Path -LiteralPath $destinationPath) {
    Remove-Item -LiteralPath $destinationPath -Recurse -Force
}
New-Item -ItemType Directory -Path $destinationPath -Force | Out-Null

foreach ($relativePath in $sourceFiles) {
    $sourcePath = Join-Path $repoRoot $relativePath
    Copy-Item -LiteralPath $sourcePath -Destination (Join-Path $destinationPath ([System.IO.Path]::GetFileName($relativePath)))
}

Write-Output "Staged 3-file import payload: $destinationPath"

[CmdletBinding()]
param(
    [string]$SourcePath = '',
    [string]$TargetPath = '_inspect/gui_verify_target.xls',
    [string]$ReportPath = '_inspect/gui_verify_report.html'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Convert-ToWorkspacePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [switch]$MustExist
    )

    $candidate = if ([System.IO.Path]::IsPathRooted($Path)) {
        $Path
    }
    else {
        Join-Path $PSScriptRoot $Path
    }

    $fullPath = [System.IO.Path]::GetFullPath($candidate)
    if ($MustExist) {
        return (Resolve-Path -LiteralPath $fullPath).Path
    }

    return $fullPath
}

if ([string]::IsNullOrWhiteSpace($SourcePath)) {
    $defaultSource = Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot 'old') -Filter '*_temp.xls' |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    if ($null -eq $defaultSource) {
        throw 'No *_temp.xls file was found under old. Specify -SourcePath explicitly.'
    }

    $sourceFullPath = $defaultSource.FullName
}
else {
    $sourceFullPath = Convert-ToWorkspacePath -Path $SourcePath -MustExist
}

$targetFullPath = Convert-ToWorkspacePath -Path $TargetPath
$reportFullPath = Convert-ToWorkspacePath -Path $ReportPath
$configFullPath = Convert-ToWorkspacePath -Path 'exe/generate_akanecco_shift_gui/akanecco_shift_config.json' -MustExist

$targetDirectory = Split-Path -Parent $targetFullPath
$reportDirectory = Split-Path -Parent $reportFullPath
if ($targetDirectory) {
    New-Item -ItemType Directory -Path $targetDirectory -Force | Out-Null
}
if ($reportDirectory) {
    New-Item -ItemType Directory -Path $reportDirectory -Force | Out-Null
}

Copy-Item -LiteralPath $sourceFullPath -Destination $targetFullPath -Force

if (Test-Path -LiteralPath $reportFullPath) {
    Remove-Item -LiteralPath $reportFullPath -Force
}

$defaultAdjacentReport = Join-Path (Split-Path -Parent $targetFullPath) (([System.IO.Path]::GetFileNameWithoutExtension($targetFullPath)) + '_validation.html')
if ($defaultAdjacentReport -ne $reportFullPath -and (Test-Path -LiteralPath $defaultAdjacentReport)) {
    Remove-Item -LiteralPath $defaultAdjacentReport -Force
}

Write-Host 'Prepared GUI verification files.'
Write-Host ('Source file: ' + $sourceFullPath)
Write-Host ('Verify target: ' + $targetFullPath)
Write-Host ('Report path: ' + $reportFullPath)
Write-Host ('Config json: ' + $configFullPath)
Write-Host 'Recommended: turn off Excel auto-open during GUI verification.'

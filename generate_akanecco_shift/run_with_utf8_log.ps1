param(
    [Parameter(Mandatory = $true)]
    [string]$ExePath,

    [Parameter(Mandatory = $true)]
    [string]$LogPath,

    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Arguments
)

$utf8 = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = $utf8
$OutputEncoding = $utf8

$logDir = Split-Path -Parent $LogPath
if ($logDir -and -not (Test-Path -LiteralPath $logDir)) {
    New-Item -ItemType Directory -Path $logDir | Out-Null
}

$quotedArgs = @($Arguments | ForEach-Object {
    if ($_ -match '[\s"]') {
        '"' + ($_ -replace '"', '""') + '"'
    }
    else {
        $_
    }
})

$commandLine = '"' + $ExePath + '"'
if ($quotedArgs.Count -gt 0) {
    $commandLine += ' ' + ($quotedArgs -join ' ')
}

$outputLines = New-Object System.Collections.Generic.List[string]
$outputLines.Add(('[{0}] COMMAND: {1}' -f (Get-Date -Format 'yyyy/MM/dd HH:mm:ss.fff'), $commandLine))
$outputLines.Add('')

& $ExePath @Arguments 2>&1 | ForEach-Object {
    $line = if ($_ -is [System.Management.Automation.ErrorRecord]) { $_.ToString() } else { [string]$_ }
    $outputLines.Add($line)
    Write-Host $line
}

$exitCode = if ($null -ne $LASTEXITCODE) { [int]$LASTEXITCODE } else { 0 }
[System.IO.File]::WriteAllLines($LogPath, $outputLines, $utf8)
exit $exitCode
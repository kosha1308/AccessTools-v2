param([Parameter(Mandatory=$true)][string]$PackageDirectory, [ValidateRange(20,180)][int]$TimeoutSeconds = 60)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$runDirectory = $null; $worker = $null
try {
    $PackageDirectory = (Resolve-Path -LiteralPath $PackageDirectory).Path
    $manifest = Get-Content -LiteralPath (Join-Path $PackageDirectory 'manifest.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $actualFiles = @(Get-ChildItem -LiteralPath $PackageDirectory -File -Recurse)
    if ($actualFiles.Count -ne $manifest.Files.Count + 1) { throw 'PILOT_MANIFEST_FILE_COUNT' }
    $seen = @{}
    foreach ($item in $manifest.Files) {
        $target = [IO.Path]::GetFullPath((Join-Path $PackageDirectory $item.Path))
        if (-not $target.StartsWith($PackageDirectory + '\', [StringComparison]::OrdinalIgnoreCase) -or $seen.ContainsKey($target)) { throw 'PILOT_MANIFEST_PATH' }
        $seen[$target] = $true
        if ((Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash -cne $item.SHA256) { throw 'PILOT_MANIFEST_HASH' }
    }
    'Manifest SHA256 PASS'
    $root = Split-Path $PSScriptRoot -Parent
    $runDirectory = Join-Path $root ('artifacts/pilot-smoke-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $runDirectory | Out-Null
    $arguments = '-NoProfile -ExecutionPolicy Bypass -File "{0}" -PackageDirectory "{1}" -RunDirectory "{2}"' -f (Join-Path $PSScriptRoot 'PilotSmokeWorker.ps1'),$PackageDirectory,$runDirectory
    $worker = Start-Process powershell.exe -ArgumentList $arguments -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $runDirectory 'stdout.log') -RedirectStandardError (Join-Path $runDirectory 'stderr.log')
    $handle = $worker.Handle
    $finished = $worker.WaitForExit($TimeoutSeconds * 1000)
    if (-not $finished) { $worker.Kill(); $worker.WaitForExit() }
    Get-Content -LiteralPath (Join-Path $runDirectory 'stdout.log')
    Get-Content -LiteralPath (Join-Path $runDirectory 'stderr.log')
    if (-not $finished) { throw 'PILOT_SMOKE_TIMEOUT' }
    if ($worker.ExitCode -ne 0) { throw 'PILOT_SMOKE_FAILED' }
    "Pilot smoke report: $runDirectory"
} catch {
    Write-Output ('PILOT_TEST_FAILED: ' + $_.Exception.Message)
    exit 1
} finally {
    if ($runDirectory) {
        $receipt = Join-Path $runDirectory 'excel-process.json'
        if (Test-Path -LiteralPath $receipt) {
            $identity = Get-Content -LiteralPath $receipt -Raw | ConvertFrom-Json
            $owned = Get-Process -Id $identity.Id -ErrorAction SilentlyContinue
            if ($owned -and $owned.ProcessName -eq 'EXCEL' -and $owned.StartTime.ToUniversalTime().Ticks -eq $identity.StartTicks) { Stop-Process -Id $owned.Id -Force }
        }
    }
}

param([string]$Root = (Split-Path $PSScriptRoot -Parent), [ValidateRange(5,300)][int]$TimeoutSeconds = 60)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$worker = $null
$runDirectory = $null
try {
    $Root = (Resolve-Path -LiteralPath $Root).Path
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'StaticValidate.ps1') -Root $Root
    if ($LASTEXITCODE -ne 0) { exit 1 }
    $runDirectory = Join-Path $Root ('artifacts/excel-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $runDirectory -Force | Out-Null
    $workerScript = Join-Path $PSScriptRoot 'ExcelWorker.ps1'
    $arguments = '-NoProfile -ExecutionPolicy Bypass -File "{0}" -Root "{1}" -RunDirectory "{2}"' -f $workerScript, $Root, $runDirectory
    $worker = Start-Process powershell.exe -ArgumentList $arguments -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $runDirectory 'stdout.log') -RedirectStandardError (Join-Path $runDirectory 'stderr.log')
    $processHandle = $worker.Handle
    $finished = $worker.WaitForExit($TimeoutSeconds * 1000)
    if (-not $finished) { $worker.Kill(); $worker.WaitForExit() }
    foreach ($name in @('stdout.log', 'stderr.log')) {
        $log = Join-Path $runDirectory $name
        if (Test-Path -LiteralPath $log) { Get-Content -LiteralPath $log }
    }
    if (-not $finished) { throw "EXCEL_TIMEOUT: compile/test exceeded $TimeoutSeconds seconds" }
    if ($worker.ExitCode -ne 0) { throw "EXCEL_GATE_FAILED: worker exit $($worker.ExitCode)" }
    if (-not (Test-Path -LiteralPath (Join-Path $runDirectory 'vba-tests.json'))) { throw 'EXCEL_TEST_RESULT: no result artifact' }
    Write-Output "Report: $runDirectory"
} catch {
    Write-Output $_.Exception.Message
    exit 1
} finally {
    if ($runDirectory) {
        $receipt = Join-Path $runDirectory 'excel-process.json'
        if (Test-Path -LiteralPath $receipt) {
            $identity = Get-Content -Raw -LiteralPath $receipt | ConvertFrom-Json
            $owned = Get-Process -Id $identity.Id -ErrorAction SilentlyContinue
            if ($owned -and $owned.ProcessName -eq 'EXCEL' -and $owned.StartTime.ToUniversalTime().Ticks -eq $identity.StartTicks) {
                Stop-Process -Id $owned.Id -Force
            }
        }
    }
}

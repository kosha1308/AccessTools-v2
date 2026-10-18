param([string]$Root = '', [Parameter(Mandatory=$true)][string]$OutputDirectory,
    [ValidatePattern('^[0-9A-Za-z.-]{1,50}$')][string]$Version = '2.0.0-pilot.5', [ValidateRange(30,300)][int]$TimeoutSeconds = 90)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$worker = $null; $runDirectory = $null
try {
    if (-not $Root) { $Root = Split-Path $PSScriptRoot -Parent }
    $Root = (Resolve-Path -LiteralPath $Root).Path
    $OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
    if (Test-Path -LiteralPath $OutputDirectory) { throw 'PILOT_OUTPUT_EXISTS' }
    $dirty = @(& git -C $Root status --porcelain)
    if ($LASTEXITCODE -ne 0 -or $dirty.Count -gt 0) { throw 'PILOT_REQUIRES_CLEAN_COMMIT' }
    $commit = [string](& git -C $Root rev-parse HEAD)
    if ($LASTEXITCODE -ne 0 -or $commit -notmatch '^[0-9a-f]{40}$') { throw 'PILOT_COMMIT_INVALID' }
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $Root 'tools/Test.ps1') -Root $Root -TimeoutSeconds $TimeoutSeconds
    if ($LASTEXITCODE -ne 0) { throw 'PILOT_REGRESSION_FAILED' }
    # No previous output is removed, even when a build fails.
    New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
    $package = Join-Path $OutputDirectory 'AccessTools'
    $runDirectory = Join-Path $OutputDirectory 'build-diagnostics'
    New-Item -ItemType Directory -Path $package, $runDirectory, (Join-Path $package 'config') | Out-Null
    $releaseId = "$Version-$($commit.Substring(0,12))"
    $ascii = [Text.ASCIIEncoding]::new()
    [IO.File]::WriteAllText((Join-Path $package 'VERSION.txt'), $Version, $ascii)
    [IO.File]::WriteAllText((Join-Path $package 'RELEASE_INFO.txt'), "Version=$Version`nReleaseId=$releaseId`nRuntimeProfile=COORDINATOR`nCommit=$commit", $ascii)
    [IO.File]::WriteAllText((Join-Path $package 'config/runtime.ini'), "SchemaVersion=1.0.0`nRuntimeProfile=COORDINATOR", $ascii)
    Copy-Item -LiteralPath (Join-Path $Root 'resources/demo') -Destination (Join-Path $package 'Пример') -Recurse
    Copy-Item -LiteralPath (Join-Path $Root 'resources/pilot/Прочитайте-сначала.txt') -Destination $package
    $workerScript = Join-Path $Root 'tools/PilotBuildWorker.ps1'
    $arguments = '-NoProfile -ExecutionPolicy Bypass -File "{0}" -Root "{1}" -RunDirectory "{2}" -PackageDirectory "{3}" -Version "{4}" -ReleaseId "{5}" -Commit "{6}"' -f $workerScript,$Root,$runDirectory,$package,$Version,$releaseId,$commit
    $worker = Start-Process powershell.exe -ArgumentList $arguments -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $runDirectory 'stdout.log') -RedirectStandardError (Join-Path $runDirectory 'stderr.log')
    $handle = $worker.Handle
    $finished = $worker.WaitForExit($TimeoutSeconds * 1000)
    if (-not $finished) { $worker.Kill(); $worker.WaitForExit() }
    Get-Content -LiteralPath (Join-Path $runDirectory 'stdout.log')
    Get-Content -LiteralPath (Join-Path $runDirectory 'stderr.log')
    if (-not $finished) { throw 'PILOT_BUILD_TIMEOUT' }
    if ($worker.ExitCode -ne 0) { throw 'PILOT_BUILD_FAILED' }
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $Root 'tools/AddRibbon.ps1') -Path (Join-Path $package 'AccessToolsv2.xlam') -RibbonPath (Join-Path $Root 'resources/ribbon/customUI.xml')
    if ($LASTEXITCODE -ne 0) { throw 'PILOT_RIBBON_FAILED' }
    $files = @(Get-ChildItem -LiteralPath $package -Recurse -File | Sort-Object FullName | ForEach-Object {
        @{ Path=$_.FullName.Substring($package.Length+1).Replace('\','/'); SHA256=(Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash }
    })
    @{ Version=$Version; ReleaseId=$releaseId; Profile='COORDINATOR'; Commit=$commit; Files=$files } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $package 'manifest.json') -Encoding UTF8
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $Root 'tools/TestPilot.ps1') -PackageDirectory $package
    if ($LASTEXITCODE -ne 0) { throw 'PILOT_INSTALLED_SMOKE_FAILED' }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    Add-Type -AssemblyName System.IO.Compression
    $zipPath = Join-Path $OutputDirectory "AccessTools_Coordinator_$Version.zip"
    $archive = [IO.Compression.ZipFile]::Open($zipPath, [IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($file in Get-ChildItem -LiteralPath $package -Recurse -File) {
            $relative = 'AccessTools/' + $file.FullName.Substring($package.Length+1).Replace('\','/')
            [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, $file.FullName, $relative)
        }
    } finally { $archive.Dispose() }
    Get-FileHash -LiteralPath $zipPath | Format-List
    Write-Output "Pilot package: $zipPath"
} catch {
    Write-Output ('PILOT_FAILED: ' + $_.Exception.Message)
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

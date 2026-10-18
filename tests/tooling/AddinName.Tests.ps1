Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repo = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
foreach ($path in @('tools/BuildPilot.ps1','tools/PilotBuildWorker.ps1','tools/PilotSmokeWorker.ps1')) {
    $text = [IO.File]::ReadAllText((Join-Path $repo $path))
    if (-not $text.Contains("'AccessToolsv2.xlam'")) { throw "FAIL v2 addin filename missing: $path" }
}
'PASS build/save/reload/smoke use v2 filename'
$guide = [IO.File]::ReadAllText((Join-Path $repo 'resources/pilot/Прочитайте-сначала.txt'))
if (-not $guide.Contains('AccessToolsv2.xlam')) { throw 'FAIL installation guide does not use v2 filename' }
'PASS installation guide uses v2 filename'
'Addin name tooling: 2 PASS / 0 FAIL'

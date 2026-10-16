param([string]$OutputDirectory)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repo = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $repo ('artifacts/pilot-build-' + [guid]::NewGuid().ToString('N')) }
if (-not (Test-Path -LiteralPath (Join-Path $repo 'tools/BuildPilot.ps1'))) { throw 'FAIL pilot builder is not implemented' }
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repo 'tools/BuildPilot.ps1') -Root $repo -OutputDirectory $OutputDirectory
if ($LASTEXITCODE -ne 0) { throw 'FAIL pilot build' }
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = @(Get-ChildItem -LiteralPath $OutputDirectory -Filter '*.zip' -File)
if ($zip.Count -ne 1) { throw 'FAIL exactly one pilot ZIP required' }
$unpacked = Join-Path $OutputDirectory 'extracted-install'
[IO.Compression.ZipFile]::ExtractToDirectory($zip[0].FullName, $unpacked)
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repo 'tools/TestPilot.ps1') -PackageDirectory (Join-Path $unpacked 'AccessTools')
if ($LASTEXITCODE -ne 0) { throw 'FAIL packaged addin smoke test' }
'Pilot build/install integration PASS'

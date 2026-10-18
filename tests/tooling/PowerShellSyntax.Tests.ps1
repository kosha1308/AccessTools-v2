Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repo = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$passed = 0
foreach ($file in Get-ChildItem -LiteralPath (Join-Path $repo 'tools') -Filter '*.ps1' -File) {
    $tokens = $null; $parseErrors = $null
    [void][Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count -gt 0) { throw ('FAIL ' + $file.Name + ': ' + ($parseErrors.Message -join '; ')) }
    $passed++
}
Write-Output "PowerShell syntax: $passed PASS / 0 FAIL"

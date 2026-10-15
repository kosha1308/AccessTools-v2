Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repo = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$scratch = Join-Path $repo ('artifacts/static-fixtures-' + [guid]::NewGuid().ToString('N'))
$validator = Join-Path $repo 'tools/StaticValidate.ps1'
$cases = @(
    @{ Name='empty'; Files=@{}; Accept=$false },
    @{ Name='missing'; Files=@{'src/a.bas'='Option Explicit'}; Accept=$false },
    @{ Name='long'; Files=@{'src/a.bas'='Attribute VB_Name = "ABCDEFGHIJKLMNOPQRSTUVWXYZ123456"'}; Accept=$false },
    @{ Name='invalid'; Files=@{'src/a.bas'='Attribute VB_Name = "1bad"'}; Accept=$false },
    @{ Name='duplicate'; Files=@{'src/a.bas'='Attribute VB_Name = "same"'; 'tests/vba/b.bas'='Attribute VB_Name = "SAME"'}; Accept=$false },
    @{ Name='multiple'; Files=@{'src/a.bas'="Attribute VB_Name = `"One`"`nAttribute VB_Name = `"Two`""}; Accept=$false },
    @{ Name='valid'; Files=@{'src/a.bas'="Attribute VB_Name = `"GoodModule`"`nOption Explicit"; 'src/b.cls'="VERSION 1.0 CLASS`nBEGIN`n  MultiUse = -1`nEND`nAttribute VB_Name = `"GoodClass`"`nOption Explicit"}; Accept=$true }
)
$passed = 0
try {
    if (-not (Test-Path -LiteralPath $validator)) { throw 'FAIL: validator is not implemented' }
    foreach ($case in $cases) {
        $root = Join-Path $scratch $case.Name
        New-Item -ItemType Directory -Path $root -Force | Out-Null
        foreach ($entry in $case.Files.GetEnumerator()) {
            $target = Join-Path $root $entry.Key
            New-Item -ItemType Directory -Path (Split-Path $target -Parent) -Force | Out-Null
            [IO.File]::WriteAllText($target, $entry.Value)
        }
        $output = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $validator -Root $root 2>&1
        if (($LASTEXITCODE -eq 0) -ne $case.Accept) { throw "FAIL $($case.Name): $output" }
        $passed++
        Write-Output "PASS $($case.Name)"
    }
    Write-Output "Static tooling: $passed PASS / 0 FAIL"
} finally {
    # Scratch lives under artifacts and is retained for diagnostics.
}

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repo = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$scratch = Join-Path $repo ('artifacts/excel-fixtures-' + [guid]::NewGuid().ToString('N'))
$cases = @(
    @{ Name='passing'; Body='AT_RunTests = "{""passed"":1,""failed"":0,""failures"":""""}"'; Tail=''; Accept=$true },
    @{ Name='assertion-failure'; Body='AT_RunTests = "{""passed"":1,""failed"":1,""failures"":""intentional""}"'; Tail=''; Accept=$false },
    @{ Name='unexecuted-syntax-error'; Body='AT_RunTests = "{""passed"":1,""failed"":0,""failures"":""""}"'; Tail="Private Sub NeverCalled()`r`n    Dim broken As`r`nEnd Sub"; Accept=$false },
    @{ Name='empty-tests'; Body='AT_RunTests = "{""passed"":0,""failed"":0,""failures"":""""}"'; Tail=''; Accept=$false }
)
$passed = 0
foreach ($case in $cases) {
    $root = Join-Path $scratch $case.Name
    $directory = Join-Path $root 'tests/vba'
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
    $module = "Attribute VB_Name = `"modGateFixture`"`r`nOption Explicit`r`nPublic Function AT_RunTests() As String`r`n$($case.Body)`r`nEnd Function`r`n$($case.Tail)`r`n"
    [IO.File]::WriteAllText((Join-Path $directory 'modGateFixture.bas'), $module)
    $output = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repo 'tools/Test.ps1') -Root $root -TimeoutSeconds 15 2>&1
    if (($LASTEXITCODE -eq 0) -ne $case.Accept) { throw "FAIL $($case.Name): $output" }
    if ($case.Name -eq 'unexecuted-syntax-error' -and ($output -join "`n") -match '\[RUN\]') { throw 'FAIL: tests executed before successful compilation' }
    $passed++
    Write-Output "PASS $($case.Name)"
}
Write-Output "Excel gate integration: $passed PASS / 0 FAIL"

param([Parameter(Mandatory=$true)][string]$Root, [Parameter(Mandatory=$true)][string]$RunDirectory, [switch]$UpdateCatalog)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$excel = $null
$book = $null
$project = $null
$compile = $null
$exitCode = 1
try {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class ATExcelProcess {
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    [uint32]$ownedId = 0
    [void][ATExcelProcess]::GetWindowThreadProcessId([intptr]$excel.Hwnd, [ref]$ownedId)
    if ($ownedId -eq 0) { throw 'EXCEL_OWNERSHIP: cannot identify created Excel process' }
    $owned = Get-Process -Id $ownedId
    @{ Id=$ownedId; StartTicks=$owned.StartTime.ToUniversalTime().Ticks } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $RunDirectory 'excel-process.json')
    $book = $excel.Workbooks.Add()
    $project = $book.VBProject
    if ($null -eq $project) { throw 'EXCEL_VBPROJECT_ACCESS: Trust Center access unavailable' }
    $imports = @()
    foreach ($folder in @('src', 'tests/vba')) {
        $directory = Join-Path $Root $folder
        if (Test-Path -LiteralPath $directory) {
            $imports += @(Get-ChildItem -LiteralPath $directory -Recurse -File | Where-Object { $_.Extension -in @('.bas', '.cls') } | Sort-Object FullName)
        }
    }
    $importDirectory = Join-Path $RunDirectory 'import'
    New-Item -ItemType Directory -Path $importDirectory -Force | Out-Null
    foreach ($file in $imports) {
        # VBE import uses the system ANSI codepage. Reject lossy conversion.
        $text = [IO.File]::ReadAllText($file.FullName)
        $encoding = [Text.Encoding]::GetEncoding([Globalization.CultureInfo]::CurrentCulture.TextInfo.ANSICodePage, [Text.EncoderFallback]::ExceptionFallback, [Text.DecoderFallback]::ExceptionFallback)
        $bytes = $encoding.GetBytes($text.Replace("`r`n", "`n").Replace("`n", "`r`n"))
        $staged = Join-Path $importDirectory $file.Name
        if (Test-Path -LiteralPath $staged) { throw "EXCEL_IMPORT_COLLISION: $($file.Name)" }
        [IO.File]::WriteAllBytes($staged, $bytes)
        Write-Output "[IMPORT] $($file.Name)"
        $component = $project.VBComponents.Import($staged)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($component)
    }
    if (Test-Path -LiteralPath (Join-Path $Root 'src/ui/forms')) {
        & (Join-Path $Root 'tools/ImportForms.ps1') -Project $project -Root $Root
    }
    $book.Activate()
    $excel.VBE.ActiveVBProject = $project
    # VBE command 578 is Compile VBAProject. A syntax error opens a modal dialog;
    # the supervising process times out and cleans only this owned Excel instance.
    $compile = $excel.VBE.CommandBars.FindControl(1, 578)
    if ($null -eq $compile) { throw 'EXCEL_COMPILE_UNAVAILABLE: Compile VBAProject command absent' }
    Write-Output '[COMPILE] VBAProject'
    if (-not $compile.Enabled) { throw 'EXCEL_COMPILE_UNAVAILABLE: new imported project was not marked compilable' }
    $compile.Execute()
    if ($compile.Enabled) {
        try {
            $pane = $excel.VBE.ActiveCodePane
            [int]$startLine = 0; [int]$startColumn = 0; [int]$endLine = 0; [int]$endColumn = 0
            $pane.GetSelection([ref]$startLine, [ref]$startColumn, [ref]$endLine, [ref]$endColumn)
            Write-Output ("[COMPILE] Location: {0}:{1}:{2}" -f $pane.CodeModule.Parent.Name, $startLine, $startColumn)
            if ($startLine -gt 0) { Write-Output $pane.CodeModule.Lines($startLine, 1) }
        } catch { Write-Output '[COMPILE] No error location available' }
        throw 'EXCEL_COMPILE_FAILED: command remains enabled'
    }
    Write-Output '[COMPILE] PASS'
    Write-Output '[RUN] AT_RunTests'
    if (Test-Path -LiteralPath (Join-Path $Root 'tests/vba/modATTestMain.bas')) {
        $json = [string]$excel.Run("'$($book.Name)'!AT_RunTests", (Join-Path $RunDirectory 'test-stages.log'))
    } else {
        $json = [string]$excel.Run("'$($book.Name)'!AT_RunTests")
    }
    $result = $json | ConvertFrom-Json
    if (($result.passed -isnot [int] -and $result.passed -isnot [long]) -or ($result.failed -isnot [int] -and $result.failed -isnot [long])) { throw 'EXCEL_TEST_RESULT: counters must be JSON integers' }
    if ($result.passed -lt 1 -or $result.failed -lt 0 -or $result.failures -isnot [string]) { throw 'EXCEL_TEST_RESULT: invalid or empty test result' }
    [IO.File]::WriteAllText((Join-Path $RunDirectory 'vba-tests.json'), $json)
    Write-Output "VBA tests: $($result.passed) PASS / $($result.failed) FAIL"
    if ($result.failed -ne 0) { throw "EXCEL_TEST_FAILURE: $($result.failures)" }
    if (Test-Path -LiteralPath (Join-Path $Root 'src/admin/modATLanguageDocs.bas')) {
        $catalog = [string]$excel.Run("'$($book.Name)'!AT_ExportLanguageCatalog")
        if ([string]::IsNullOrWhiteSpace($catalog)) { throw 'CATALOG_EMPTY: registry documentation was empty' }
        [IO.File]::WriteAllText((Join-Path $RunDirectory 'language-catalog.md'), $catalog)
        $snapshot = Join-Path $Root 'docs/dsl/catalog.md'
        if ($UpdateCatalog) {
            [IO.File]::WriteAllText($snapshot, $catalog)
        } elseif (Test-Path -LiteralPath $snapshot) {
            $existingCatalog = [IO.File]::ReadAllText($snapshot).Replace("`r`n", "`n")
            if ($existingCatalog -cne $catalog.Replace("`r`n", "`n")) { throw 'CATALOG_DRIFT: regenerate docs/dsl/catalog.md from the emitted language-catalog.md artifact' }
        } else {
            throw 'CATALOG_MISSING: run tools/Test.ps1 -UpdateCatalog'
        }
    }
    $exitCode = 0
} catch {
    Write-Output ('EXCEL_GATE_FAILED: ' + $_.Exception.Message)
} finally {
    if ($book) { try { $book.Close($false) } catch {} }
    if ($excel) { try { $excel.Quit() } catch {} }
    foreach ($obj in @($compile, $project, $book, $excel)) {
        if ($null -ne $obj -and [Runtime.InteropServices.Marshal]::IsComObject($obj)) {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($obj)
        }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
exit $exitCode

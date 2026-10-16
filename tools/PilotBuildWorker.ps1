param([Parameter(Mandatory=$true)][string]$Root, [Parameter(Mandatory=$true)][string]$RunDirectory,
[Parameter(Mandatory=$true)][string]$PackageDirectory, [Parameter(Mandatory=$true)][ValidatePattern('^[0-9A-Za-z.-]{1,50}$')][string]$Version,
[Parameter(Mandatory=$true)][ValidatePattern('^[0-9A-Za-z.-]{1,100}$')][string]$ReleaseId,
[Parameter(Mandatory=$true)][ValidatePattern('^[0-9a-f]{40}$')][string]$Commit)
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
    foreach ($folder in @('src')) {
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
        if ($file.Name -eq 'modATPackage.bas') {
            $text = $text.Replace('AT_PACKAGE_VERSION As String = "DEV"', ('AT_PACKAGE_VERSION As String = "' + $Version + '"'))
            $text = $text.Replace('AT_PACKAGE_RELEASE As String = "DEV"', ('AT_PACKAGE_RELEASE As String = "' + $ReleaseId + '"'))
            $text = $text.Replace('AT_PACKAGE_PROFILE As String = "DEV"', 'AT_PACKAGE_PROFILE As String = "COORDINATOR"')
            $text = $text.Replace('AT_PACKAGE_COMMIT As String = "DEV"', ('AT_PACKAGE_COMMIT As String = "' + $Commit + '"'))
            if ($text -match 'As String = "DEV"') { throw 'PILOT_DEV_CONSTANT_RETAINED' }
        }
        $encoding = [Text.Encoding]::GetEncoding([Globalization.CultureInfo]::CurrentCulture.TextInfo.ANSICodePage, [Text.EncoderFallback]::ExceptionFallback, [Text.DecoderFallback]::ExceptionFallback)
        $bytes = $encoding.GetBytes($text.Replace("`r`n", "`n").Replace("`n", "`r`n"))
        $staged = Join-Path $importDirectory $file.Name
        if (Test-Path -LiteralPath $staged) { throw "EXCEL_IMPORT_COLLISION: $($file.Name)" }
        [IO.File]::WriteAllBytes($staged, $bytes)
        Write-Output "[IMPORT] $($file.Name)"
        $component = $project.VBComponents.Import($staged)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($component)
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
    $identity = [string]$excel.Run("'$($book.Name)'!AT_PackageIdentity")
    if ($identity -cne "$Version|$ReleaseId|COORDINATOR|$Commit") { throw 'PILOT_EMBEDDED_IDENTITY_INVALID' }
    $book.IsAddin = $true
    $book.SaveAs((Join-Path $PackageDirectory 'AccessTools.xlam'), 55)
    if ($book.FileFormat -ne 55 -or -not $book.IsAddin) { throw 'PILOT_FORMAT_INVALID' }
    if ([string]$excel.Run("'$($book.Name)'!AT_PackageInstalledCode") -ne '') { throw 'PILOT_METADATA_INVALID' }
    @{Identity=$identity; Modules=@($imports.Name); FileFormat=$book.FileFormat} | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath (Join-Path $RunDirectory 'build.json') -Encoding UTF8
    Write-Output '[BUILD] XLAM PASS'
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

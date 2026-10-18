param([Parameter(Mandatory=$true)][string]$PackageDirectory, [Parameter(Mandatory=$true)][string]$RunDirectory)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$excel = $null; $addin = $null; $request = $null; $copy = $null; $legacy = $null; $exitCode = 1; $passed = 0
try {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class ATPilotExcelProcess {
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@
    $installParent = Join-Path $RunDirectory 'install'
    New-Item -ItemType Directory -Path $installParent | Out-Null
    $install = Join-Path $installParent 'AccessTools'
    Copy-Item -LiteralPath $PackageDirectory -Destination $install -Recurse
    $manifest = Get-Content -LiteralPath (Join-Path $install 'manifest.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.AskToUpdateLinks = $false
    [uint32]$ownedId = 0
    [void][ATPilotExcelProcess]::GetWindowThreadProcessId([intptr]$excel.Hwnd, [ref]$ownedId)
    if ($ownedId -eq 0) { throw 'PILOT_OWNERSHIP' }
    $owned = Get-Process -Id $ownedId
    @{Id=$ownedId;StartTicks=$owned.StartTime.ToUniversalTime().Ticks} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $RunDirectory 'excel-process.json')
    # A synthetic first-version filename fixture, never the user's real add-in.
    $legacyFolder = Join-Path $RunDirectory 'legacy-fixture'
    New-Item -ItemType Directory -Path $legacyFolder | Out-Null
    $legacyPath = Join-Path $legacyFolder 'AccessTools.xlam'
    $legacy = $excel.Workbooks.Add(-4167)
    $legacy.IsAddin = $true
    $legacy.SaveAs($legacyPath,55)
    $legacy.Close($false)
    $legacy = $excel.Workbooks.Open($legacyPath,0,$true)
    $legacyHash = (Get-FileHash -LiteralPath $legacyPath).Hash
    $addin = $excel.Workbooks.Open((Join-Path $install 'AccessToolsv2.xlam'), 0, $true)
    if (-not $addin.IsAddin -or $addin.FileFormat -ne 55) { throw 'PILOT_LOADED_FORMAT' }; $passed++
    $macro = "'$($addin.Name)'!"
    $identity = [string]$excel.Run($macro + 'AT_PackageIdentity')
    if ($identity -cne "$($manifest.Version)|$($manifest.ReleaseId)|COORDINATOR|$($manifest.Commit)" -or $identity -match 'DEV') { throw 'PILOT_LOADED_IDENTITY' }; $passed++
    if ([string]$excel.Run($macro + 'AT_PackageInstalledCode') -ne '') { throw 'PILOT_LOADED_METADATA' }; $passed++
    foreach ($book in $excel.Workbooks) { if ($book.Name -ne $addin.Name -and $book.Name -ne $legacy.Name) { throw 'PILOT_STARTUP_OPENED_BOOKS' } }; $passed++
    $names = @($addin.VBProject.VBComponents | ForEach-Object { $_.Name })
    if (@($names | Where-Object { $_ -like '*Test*' }).Count -gt 0) { throw 'PILOT_TEST_MODULE_LEAK' }; $passed++
    $code = $addin.VBProject.VBComponents.Item('modATNetworkCommand').CodeModule
    $text = $code.Lines(1, $code.CountOfLines)
    if ($text -notmatch 'Public Sub AT_RibbonCheck' -or $text -notmatch 'Public Sub AT_RibbonSave') { throw 'PILOT_CALLBACK_MISSING' }; $passed++
    $versionPath = Join-Path $install 'VERSION.txt'
    [IO.File]::WriteAllText($versionPath, 'wrong-version')
    if ([string]$excel.Run($macro + 'AT_PackageInstalledCode') -ne 'PACKAGE_MISMATCH') { throw 'PILOT_MIXED_VERSION_ACCEPTED' }; $passed++
    [IO.File]::WriteAllText($versionPath, $manifest.Version)
    $demo = Join-Path $install 'Пример'
    $request = $excel.Workbooks.Open((Join-Path $demo 'Заявка-пример.xlsx'), 0, $true)
    $before = $excel.Workbooks.Count
    # Excel cannot marshal private VBA class results through Application.Run.
    # A temporary in-memory probe keeps the result inside this same VBA project.
    # ReadOnly addin is never saved; packaged modules were checked before injection.
    Write-Output '[SMOKE] injecting same-project probe'
    $probe = $addin.VBProject.VBComponents.Add(1)
    $probe.Name = 'modATPilotProbe'
    $probe.CodeModule.AddFromString(@"
Option Explicit
Public Function AT_PilotExecute(ByVal request As Worksheet, ByVal rulePath As String, ByVal sourcePath As String, ByVal outputParent As String) As String
    On Error GoTo FailedRun
    Dim result As ATNetworkRunResult
    Set result = AT_CommandExecute(request, rulePath, sourcePath, "PILOT_SMOKE", "2026-10-17", "Europe/Moscow")
    If Not result.IsComplete Then Err.Raise vbObjectError + 3190, , result.ErrorCode
    AT_PilotExecute = AT_CommandSave(outputParent)
    Exit Function
FailedRun:
    AT_PilotExecute = "FAILED: " & Err.Description
End Function
"@)
    Write-Output '[SMOKE] running same-project workflow'
    $saved = [string]$excel.Run($macro + 'AT_PilotExecute', $request.Worksheets.Item(1), (Join-Path $demo '01_Правила_NETWORK_ACCESS.xlsx'), (Join-Path $demo 'Выгрузка-Superset.xlsx'), $RunDirectory)
    Write-Output ("[SMOKE] workflow returned: {0}" -f $saved)
    if ($saved.StartsWith('FAILED:')) { throw $saved }
    if (-not (Test-Path -LiteralPath (Join-Path $saved 'Отчёт-проверок.xlsx')) -or -not (Test-Path -LiteralPath (Join-Path $saved 'Сводка-координатора.xlsx'))) { throw 'PILOT_OUTPUTS_MISSING' }; $passed++
    if ($excel.Workbooks.Count -ne $before + 2 -or -not $request.ReadOnly -or -not $request.Saved) { throw 'PILOT_BOOK_OWNERSHIP' }; $passed++
    $copy = $excel.Workbooks.Open((Join-Path $saved 'Сводка-координатора.xlsx'), 0, $true)
    if ($copy.Worksheets.Count -ne 2 -or $copy.Worksheets.Item(1).Cells.Item(6,5).Value2 -ne 'Сервис из правил' -or $copy.Worksheets.Item(1).Cells.Item(6,6).Value2 -notlike '*Техническая инструкция*') { throw 'PILOT_OUTPUT_CONTENT' }; $passed++
    $copy.Close($false); $copy = $null
    foreach ($item in $manifest.Files) {
        if ((Get-FileHash -LiteralPath (Join-Path $install $item.Path)).Hash -cne $item.SHA256) { throw 'PILOT_INPUT_CHANGED' }
    }; $passed++
    if ($addin.Name -cne 'AccessToolsv2.xlam' -or $legacy.Name -cne 'AccessTools.xlam' -or -not $legacy.IsAddin -or -not $legacy.ReadOnly -or -not $legacy.Saved -or (Get-FileHash -LiteralPath $legacyPath).Hash -cne $legacyHash) { throw 'PILOT_LEGACY_COEXISTENCE_FAILED' }; $passed++
    @{Passed=$passed;Failed=0;Identity=$identity;Saved=$saved;Modules=$names} | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath (Join-Path $RunDirectory 'smoke.json') -Encoding UTF8
    "Pilot installed XLAM: $passed PASS / 0 FAIL"
    $exitCode = 0
} catch { Write-Output ('PILOT_SMOKE_FAILED: ' + $_.Exception.Message) }
finally {
    if ($copy) { try { $copy.Close($false) } catch {} }
    if ($request) { try { $request.Close($false) } catch {} }
    if ($excel) {
        try { for ($i=$excel.Workbooks.Count; $i -ge 1; $i--) { $excel.Workbooks.Item($i).Close($false) } } catch {}
    }
    if ($addin) { try { $addin.Close($false) } catch {} }
    if ($legacy) { try { $legacy.Close($false) } catch {} }
    if ($excel) { try { $excel.Quit() } catch {} }
    foreach ($obj in @($copy,$request,$addin,$legacy,$excel)) {
        if ($null -ne $obj -and [Runtime.InteropServices.Marshal]::IsComObject($obj)) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($obj) }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
exit $exitCode

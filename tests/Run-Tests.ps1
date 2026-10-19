param([switch]$Worker, [string]$RunDirectory)
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
if (-not $Worker) {
    $run = Join-Path $repo ('artifacts/run-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $run | Out-Null
    $hostExe = Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/powershell.exe'
    $arguments = @('-NoProfile','-ExecutionPolicy','Bypass','-File',('"'+$PSCommandPath+'"'),'-Worker','-RunDirectory',('"'+$run+'"'))
    $p = Start-Process $hostExe -ArgumentList $arguments -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $run 'stdout.txt') -RedirectStandardError (Join-Path $run 'stderr.txt')
    if (-not $p.WaitForExit(90000)) {
        $ownerFile = Join-Path $run 'excel-owner.json'
        if (Test-Path $ownerFile) {
            $owner = Get-Content $ownerFile -Raw | ConvertFrom-Json
            $owned = Get-Process -Id $owner.Id -ErrorAction SilentlyContinue
            if ($owned -and $owned.ProcessName -eq 'EXCEL' -and $owned.StartTime.ToUniversalTime().Ticks.ToString() -eq $owner.StartTicks) { Stop-Process -Id $owned.Id -Force }
        }
        Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
        throw "Excel test timeout; evidence: $run"
    }
    Get-Content (Join-Path $run 'stdout.txt')
    if ($p.ExitCode -ne 0) { Get-Content (Join-Path $run 'stderr.txt'); throw "Tests failed; evidence: $run" }
    Write-Output "Evidence: $run"
    exit 0
}
$excel = $null; $book = $null
try {
    Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class AtTestWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint p); }'
    $excel = New-Object -ComObject Excel.Application
    [uint32]$ownedId = 0
    [void][AtTestWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$ownedId)
    $owned = Get-Process -Id $ownedId
    @{Id=$ownedId; StartTicks=$owned.StartTime.ToUniversalTime().Ticks.ToString()} | ConvertTo-Json | Set-Content (Join-Path $RunDirectory 'excel-owner.json')
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Add()
    $project = $book.VBProject
    $files = @()
    if (Test-Path (Join-Path $repo 'src')) { $files += Get-ChildItem (Join-Path $repo 'src') -File | Where-Object { $_.Extension -in '.bas','.cls' } }
    $files += Get-ChildItem $PSScriptRoot -Filter '*.bas'
    foreach ($file in $files) {
        $source = [IO.File]::ReadAllText($file.FullName)
        $moduleName = [regex]::Match($source,'Attribute VB_Name = "([^"]+)"').Groups[1].Value
        if (-not $moduleName) { throw "Missing module name: $file" }
        $source = ($source -split "`r?`n" | Where-Object { $_ -notmatch '^(Attribute |VERSION |BEGIN$|END$|  MultiUse)' }) -join "`r`n"
        $component = $project.VBComponents.Add($(if ($file.Extension -eq '.cls') {2} else {1}))
        $component.Name = $moduleName
        $component.CodeModule.AddFromString($source)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($component)
    }
    $book.Activate()
    "Excel=$($excel.Version); Build=$($excel.Build); OS=$($excel.OperatingSystem); DeveloperHostBits=$([IntPtr]::Size * 8)" | Set-Content (Join-Path $RunDirectory 'host.txt')
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $hashes = foreach ($file in $files) {
            [BitConverter]::ToString($sha.ComputeHash([IO.File]::ReadAllBytes($file.FullName))).Replace('-','') + ' ' + $file.Name
        }
        $hashes | Set-Content (Join-Path $RunDirectory 'source-hashes.txt')
    } finally { $sha.Dispose() }
    $result = [string]$excel.Run("'"+$book.Name+"'!AtRunSchemaTests", $repo, $RunDirectory)
    $result | Set-Content (Join-Path $RunDirectory 'result.txt')
    Write-Output "Excel $($excel.Version) | $result"
    if (-not $result.StartsWith('PASS|')) { throw $result }
    $astResult = [string]$excel.Run("'"+$book.Name+"'!AtRunAstTests", $repo, $RunDirectory)
    $astResult | Set-Content (Join-Path $RunDirectory 'ast-result.txt')
    Write-Output $astResult
    if (-not $astResult.StartsWith('PASS|')) { throw $astResult }
    $evalResult = [string]$excel.Run("'"+$book.Name+"'!AtRunEvalTests", $repo, $RunDirectory)
    $evalResult | Set-Content (Join-Path $RunDirectory 'eval-result.txt')
    Write-Output $evalResult
    if (-not $evalResult.StartsWith('PASS|')) { throw $evalResult }
    $resolverResult = [string]$excel.Run("'"+$book.Name+"'!AtRunResolverTests", $repo, $RunDirectory)
    $resolverResult | Set-Content (Join-Path $RunDirectory 'resolver-result.txt')
    Write-Output $resolverResult
    if (-not $resolverResult.StartsWith('PASS|')) { throw $resolverResult }
    $writerResult = [string]$excel.Run("'"+$book.Name+"'!AtRunWriterTests", $repo, $RunDirectory)
    $writerResult | Set-Content (Join-Path $RunDirectory 'writer-result.txt')
    Write-Output $writerResult
    if (-not $writerResult.StartsWith('PASS|')) { throw $writerResult }
} finally {
    if ($book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
}

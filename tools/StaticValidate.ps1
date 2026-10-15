param([string]$Root = (Split-Path $PSScriptRoot -Parent))
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
try {
    $files = @()
    foreach ($folder in @('src', 'tests/vba')) {
        $directory = Join-Path $Root $folder
        if (Test-Path -LiteralPath $directory) {
            $files += @(Get-ChildItem -LiteralPath $directory -Recurse -File | Where-Object { $_.Extension -in @('.bas', '.cls') })
        }
    }
    if ($files.Count -eq 0) { throw 'STATIC_NO_SOURCES: no VBA modules found' }
    $seen = @{}
    foreach ($file in $files) {
        $source = [IO.File]::ReadAllText($file.FullName)
        $names = [regex]::Matches($source, '(?m)^Attribute VB_Name = "([^"]+)"\r?$')
        if ($names.Count -ne 1) { throw "STATIC_MODULE_NAME: exactly one VB_Name required: $($file.FullName)" }
        $name = $names[0].Groups[1].Value
        if ($name.Length -gt 31 -or $name -cnotmatch '^[A-Za-z][A-Za-z0-9_]*$') { throw "STATIC_MODULE_NAME: invalid component name: $name" }
        if ($seen.ContainsKey($name)) { throw "STATIC_DUPLICATE_NAME: $name in $($file.FullName) and $($seen[$name])" }
        $seen[$name] = $file.FullName
    }
    Write-Output "Static validation: $($files.Count) components PASS"
    exit 0
} catch {
    Write-Output $_.Exception.Message
    exit 1
}

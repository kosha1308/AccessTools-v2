param([switch]$Worker)
$ErrorActionPreference = 'Stop'
if (-not $Worker) {
    foreach ($hostPath in @('System32','SysWOW64')) {
        $exe = Join-Path $env:WINDIR "$hostPath/WindowsPowerShell/v1.0/powershell.exe"
        & $exe -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath -Worker
        if ($LASTEXITCODE -ne 0) { throw "MSXML probe failed: $hostPath" }
    }
    exit 0
}
$dom = New-Object -ComObject Msxml2.DOMDocument.6.0
try {
    $dom.async = $false; $dom.resolveExternals = $false
    $dom.setProperty('ProhibitDTD', $true)
    $dom.setProperty('MaxXMLSize', 1024)
    $dom.setProperty('MaxElementDepth', 32)
    $text = '<root>' + [char]0x418 + [char]0x43C + [char]0x44F + '</root>'
    if (-not $dom.loadXML($text)) { throw 'Unicode parsing failed' }
    if ($dom.documentElement.text -ne ($text -replace '</?root>','')) { throw 'Unicode mismatch' }
    if ($dom.loadXML('<!DOCTYPE root [<!ENTITY x "blocked">]><root>&x;</root>')) { throw 'DTD was accepted' }
    Write-Output "PASS MSXML6 $([IntPtr]::Size * 8)-bit: activation, Unicode, secure properties, DTD rejection"
} finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($dom) }

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression
$repo = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$scratch = Join-Path $repo ('artifacts/ribbon-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $scratch | Out-Null
$fixture = Join-Path $scratch 'fixture'
New-Item -ItemType Directory -Path (Join-Path $fixture '_rels') -Force | Out-Null
[IO.File]::WriteAllText((Join-Path $fixture '_rels/.rels'), '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="existing" Type="urn:keep" Target="keep.bin"/></Relationships>')
[IO.File]::WriteAllText((Join-Path $fixture 'keep.bin'), 'keep unchanged')
$book = Join-Path $scratch 'test.xlam'
$zip = [IO.Compression.ZipFile]::Open($book, [IO.Compression.ZipArchiveMode]::Create)
try {
    foreach ($file in Get-ChildItem -LiteralPath $fixture -File -Recurse) {
        $relative = $file.FullName.Substring($fixture.Length + 1).Replace('\', '/')
        [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $file.FullName, $relative)
    }
} finally { $zip.Dispose() }
$tool = Join-Path $repo 'tools/AddRibbon.ps1'
if (-not (Test-Path -LiteralPath $tool)) { throw 'FAIL Ribbon injection is not implemented' }
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path $book -RibbonPath (Join-Path $repo 'resources/ribbon/customUI.xml')
if ($LASTEXITCODE -ne 0) { throw 'FAIL first injection' }
$archive = [IO.Compression.ZipFile]::OpenRead($book)
try {
    $reader = [IO.StreamReader]::new($archive.GetEntry('keep.bin').Open())
    try { if ($reader.ReadToEnd() -cne 'keep unchanged') { throw 'FAIL existing binary' } } finally { $reader.Dispose() }
    $reader = [IO.StreamReader]::new($archive.GetEntry('_rels/.rels').Open())
    try { [xml]$rels = $reader.ReadToEnd() } finally { $reader.Dispose() }
    if ($rels.Relationships.Relationship.Count -ne 2) { throw 'FAIL relationship preservation' }
    $reader = [IO.StreamReader]::new($archive.GetEntry('customUI/customUI.xml').Open())
    try { [xml]$xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
    $buttons = @($xml.SelectNodes('//*[local-name()="button"]'))
    if ($buttons.Count -ne 9) { throw 'FAIL buttons' }
    if (@($buttons | Where-Object { $_.onAction -eq 'AT_RibbonDiagnostic' }).Count -ne 1) { throw 'FAIL diagnostic action' }
    if (@($xml.SelectNodes('//*[local-name()="group"]')).Count -ne 3) { throw 'FAIL task groups' }
} finally { $archive.Dispose() }
'PASS Ribbon part, relationships, existing binary and grouped template/result/diagnostic controls'
$hash = (Get-FileHash -LiteralPath $book).Hash
$ErrorActionPreference = 'Continue'
$output = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path $book -RibbonPath (Join-Path $repo 'resources/ribbon/customUI.xml') 2>&1
$exitCode = $LASTEXITCODE
$ErrorActionPreference = 'Stop'
if ($exitCode -eq 0 -or (Get-FileHash -LiteralPath $book).Hash -ne $hash) { throw "FAIL duplicate must reject without modification: $output" }
'PASS duplicate rejects without mutation'
'Ribbon tooling: 2 PASS / 0 FAIL'

# The expected rejected native command must not become GitHub's step exit code.
exit 0

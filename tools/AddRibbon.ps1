param([Parameter(Mandatory=$true)][string]$Path, [Parameter(Mandatory=$true)][string]$RibbonPath)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression
$Path = (Resolve-Path -LiteralPath $Path).Path
$ribbonText = [IO.File]::ReadAllText((Resolve-Path -LiteralPath $RibbonPath).Path)
[xml]$ribbon = $ribbonText
if ($ribbon.DocumentElement.LocalName -ne 'customUI' -or $ribbon.DocumentElement.NamespaceURI -ne 'http://schemas.microsoft.com/office/2006/01/customui') { throw 'RIBBON_XML_INVALID' }
$type = 'http://schemas.microsoft.com/office/2006/relationships/ui/extensibility'
# Validate in read mode first: rejected inputs stay byte-for-byte unchanged.
$archive = [IO.Compression.ZipFile]::OpenRead($Path)
try {
    if ($null -ne $archive.GetEntry('customUI/customUI.xml')) { throw 'RIBBON_ALREADY_PRESENT' }
    $entry = $archive.GetEntry('_rels/.rels')
    if ($null -eq $entry) { throw 'RIBBON_RELATIONSHIPS_MISSING' }
    $reader = [IO.StreamReader]::new($entry.Open())
    try { [xml]$rels = $reader.ReadToEnd() } finally { $reader.Dispose() }
    if ($rels.DocumentElement.NamespaceURI -ne 'http://schemas.openxmlformats.org/package/2006/relationships') { throw 'RIBBON_RELATIONSHIPS_INVALID' }
    foreach ($relationship in $rels.DocumentElement.ChildNodes) {
        if ($relationship.GetAttribute('Type') -like '*relationships/ui/extensibility') { throw 'RIBBON_ALREADY_PRESENT' }
    }
} finally { $archive.Dispose() }
$node = $rels.CreateElement('Relationship', $rels.DocumentElement.NamespaceURI)
$node.SetAttribute('Id', 'atRibbon' + [guid]::NewGuid().ToString('N'))
$node.SetAttribute('Type', $type)
$node.SetAttribute('Target', 'customUI/customUI.xml')
[void]$rels.DocumentElement.AppendChild($node)
$archive = [IO.Compression.ZipFile]::Open($Path, [IO.Compression.ZipArchiveMode]::Update)
try {
    $archive.GetEntry('_rels/.rels').Delete()
    foreach ($item in @(@{Name='_rels/.rels';Text=$rels.OuterXml}, @{Name='customUI/customUI.xml';Text=$ribbonText})) {
        $entry = $archive.CreateEntry($item.Name)
        $writer = [IO.StreamWriter]::new($entry.Open(), [Text.UTF8Encoding]::new($false))
        try { $writer.Write($item.Text) } finally { $writer.Dispose() }
    }
} finally { $archive.Dispose() }
'Ribbon injection PASS'

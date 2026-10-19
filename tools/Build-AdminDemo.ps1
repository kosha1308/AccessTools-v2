param([Parameter(Mandatory=$true)][string]$OutputPath)
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
$destination = [IO.Path]::GetFullPath($OutputPath)
if ([IO.Path]::GetExtension($destination) -ne '.xlsm') { throw 'Choose an .xlsm destination.' }
if (Test-Path -LiteralPath $destination) { throw 'Destination exists; choose a new file.' }
if (-not (Test-Path -LiteralPath (Split-Path $destination -Parent) -PathType Container)) { throw 'Parent directory must exist.' }
$excel = $null; $book = $null
try {
    # Own Excel process and synthetic workbook only; no existing user workbook is attached.
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Add()
    $project = $book.VBProject
    foreach ($file in Get-ChildItem (Join-Path $repo 'src') -File | Where-Object { $_.Extension -in '.bas','.cls' }) {
        $source = [IO.File]::ReadAllText($file.FullName)
        $name = [regex]::Match($source,'Attribute VB_Name = "([^"]+)"').Groups[1].Value
        if (-not $name) { throw "Missing module name: $file" }
        $source = ($source -split "`r?`n" | Where-Object { $_ -notmatch '^(Attribute |VERSION |BEGIN$|END$|  MultiUse)' }) -join "`r`n"
        $component = $project.VBComponents.Add($(if ($file.Extension -eq '.cls') {2} else {1}))
        $component.Name = $name; $component.CodeModule.AddFromString($source)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($component)
    }
    $sheet = $book.Worksheets(1); $sheet.Name = 'AccessTools Admin'
    $sheet.Range('A1').Value2 = 'AccessTools: Text equality editor'
    $sheet.Range('A3').Value2 = 'Open a synthetic target workbook, then click Open editor.'
    $sheet.Range('A4').Value2 = 'NEW -> SAVE -> CLOSE -> LOAD -> RUN. XML/AST editing is not required.'
    $sheet.Range('A5').Value2 = '.xlsx supports fill only; Notes require the tested .xls profile.'
    $sheet.Range('A6').Value2 = 'Correct inputs before RUN. Changed annotation/structure or inputs changed after planning block writes.'
    $sheet.Range('A7').Value2 = 'Keep this macro host open until owned marks are cleared. No automatic recovery after session loss.'
    $sheet.Columns.Item(1).ColumnWidth = 110
    $button = $sheet.Buttons().Add(20,150,180,35)
    $button.Caption = 'Open editor'
    $button.OnAction = "'" + [IO.Path]::GetFileName($destination).Replace("'","''") + "'!AtAdminOpen"
    $book.SaveAs($destination, 52)
    Write-Output "Created: $destination"
} finally {
    if ($book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
}

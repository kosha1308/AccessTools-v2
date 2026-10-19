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
    . (Join-Path $PSScriptRoot 'Import-AtProject.ps1')
    $null = Import-AtProject $project $repo
    $sheet = $book.Worksheets(1); $sheet.Name = 'AccessTools'
    $sheet.Range('A1').Value2 = 'AccessTools 2 — конструктор проверок'
    $sheet.Range('A3').Value2 = 'Откройте другую книгу Excel с тестовыми данными. Нажмите «Открыть конструктор».'
    $sheet.Range('A4').Value2 = 'Настройте поле и условие справа. Сначала используйте «Протестировать» — книга не изменится.'
    $sheet.Range('A5').Value2 = '«Нанести отметки» — отдельное действие. В .xlsx доступны только заливки, без примечаний.'
    $sheet.Range('A6').Value2 = 'Сохраняйте новые редакции в новые файлы. Не закрывайте эту книгу до очистки своих отметок.'
    $sheet.Range('A7').Value2 = 'Восстановление журнала после потери сессии автоматически не выполняется.'
    $sheet.Range('A1').Font.Size = 20
    $sheet.Columns.Item(1).ColumnWidth = 120
    $sheet.Rows.Item(1).RowHeight = 36
    $button = $sheet.Buttons().Add(20,160,250,40)
    $button.Caption = 'Открыть конструктор'
    $book.SaveAs($destination, 52)
    # SaveAs can rewrite a pre-save OnAction. Bind using the ACTUAL final name.
    $expectedAction = "'" + $book.Name.Replace("'","''") + "'!AtAdminOpen"
    $button.OnAction = $expectedAction
    $book.Save()
    if ($button.OnAction -ne $expectedAction) { throw 'Launch binding mismatch.' }
    foreach ($name in 'AtAdminMain','AtAdminResults') {
        if ($project.VBComponents.Item($name).Type -ne 3) { throw "Missing UserForm: $name" }
    }
    $smoke = [string]$excel.Run(("'" + $book.Name.Replace("'","''") + "'!AtAdminUiSmoke"))
    if (-not $smoke.StartsWith('READY|')) { throw "Form initialization failed: $smoke" }
    Write-Output $smoke
    Write-Output "Created: $destination"
} finally {
    if ($book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
}

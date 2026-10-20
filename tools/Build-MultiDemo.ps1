param([Parameter(Mandatory=$true)][string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
$destination = [IO.Path]::GetFullPath($OutputDirectory)
if (-not (Test-Path -LiteralPath $destination -PathType Container)) { throw 'Output directory must exist.' }
$demo = Join-Path $destination 'AccessTools-Constructor-PR08.xlsm'
$sample = Join-Path $destination 'PR08-synthetic.xlsx'
$config = Join-Path $destination 'PR08-example.xml'
foreach ($path in @($demo,$sample,$config)) { if (Test-Path -LiteralPath $path) { throw "Destination exists: $path" } }
& (Join-Path $PSScriptRoot 'Build-AdminDemo.ps1') -OutputPath $demo
$excel = $null; $book = $null; $target = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.AutomationSecurity = 3
    # Enable only our just-built local synthetic project in this owned process.
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($demo)
    $excel.AutomationSecurity = 3
    # Synthetic demonstration only. This module is embedded in the demo, not the product.
    $component = $book.VBProject.VBComponents.Add(1); $component.Name = 'AtSyntheticDemo'
    $component.CodeModule.AddFromString(@'
Option Explicit
Public Function AtBuildExample(ByVal path As String) As String
    Dim editor As CAtTemplateEditor, field As CAtTemplateField, rule As CAtTemplateRule
    Set editor = AtNewTemplateEditor("", False): editor.TemplateName = "Демонстрация: несколько правил"
    Set field = editor.AddField("Статус", "Input", "B2")
    Set rule = editor.AddRule(field.Id, "Ожидается good")
    rule.ExpectedText = "good": rule.Message = "Укажите good.": rule.Severity = "WARNING": rule.Color = 65535
    Set rule = editor.AddRule(field.Id, "Запрещено bad")
    rule.Condition = "NOT_EQUAL": rule.ExpectedText = "bad": rule.Message = "Значение bad запрещено.": rule.Color = 255
    Set field = editor.AddField("Оборудование", "Input", "C2")
    Set rule = editor.AddRule(field.Id, "Нужно оборудование")
    rule.Condition = "REQUIRED": rule.Message = "Заполните оборудование."
    editor.SaveFile path
    ThisWorkbook.Worksheets("AccessTools").Range("A20").Value2 = AtToXml(editor.Model)
    AtBuildExample = "READY|2|3"
End Function
Public Sub AtOpenExample()
    On Error GoTo Failed
    Dim registry As New CAtOperationRegistry, issues As Collection, package As CAtPackage
    Set package = AtFromXml(CStr(ThisWorkbook.Worksheets("AccessTools").Range("A20").Value2), registry.Catalog, issues)
    If package Is Nothing Then Err.Raise 5, , "Встроенный пример не прошёл проверку."
    AtUiOpenTemplate AtReadTemplate(package)
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Пример AccessTools"
End Sub
'@)
    $result = [string]$excel.Run(("'" + $book.Name + "'!AtBuildExample"), $config)
    if ($result -ne 'READY|2|3') { throw 'Example generation failed.' }
    $sheet = $book.Worksheets('AccessTools')
    $sheet.Rows.Item(20).Hidden = $true
    $sheet.Range('A3').Value2 = 'Откройте PR08-synthetic.xlsx в ЭТОМ ЖЕ Excel. Нажмите «Открыть пример» и явно выберите эту книгу.'
    $sheet.Range('A4').Value2 = 'B2: bad → две ошибки и красная заливка; other → одна и жёлтая; good → PASS и исходная заливка.'
    $sheet.Range('A5').Value2 = '«Протестировать» обновляет все отметки; «Только выбранное» — только чтение. Книги не сохраняются автоматически.'
    $button = $sheet.Buttons().Add(290,200,250,40)
    $button.Caption = 'Открыть пример'; $button.OnAction = "'" + $book.Name + "'!AtOpenExample"
    $book.Save()
    $target = $excel.Workbooks.Add(); $inputSheet = $target.Worksheets(1); $inputSheet.Name = 'Input'
    $inputSheet.Range('B1').Value2 = 'Статус'; $inputSheet.Range('C1').Value2 = 'Оборудование'
    $inputSheet.Range('B2').Value2 = 'bad'; $inputSheet.Range('C2').Value2 = 'Laptop'
    $inputSheet.Range('B2').Interior.Color = 15790320
    $inputSheet.Range('E2').Formula = '=1+2'
    $inputSheet.Columns.Item(2).ColumnWidth = 22; $inputSheet.Columns.Item(3).ColumnWidth = 22
    $target.SaveAs($sample,51)
    Write-Output "Created: $demo; $sample; $config"
} finally {
    if ($target) { $target.Close($false) }
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
}

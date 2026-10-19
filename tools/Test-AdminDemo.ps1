param([Parameter(Mandatory=$true)][string]$Path)
$ErrorActionPreference = 'Stop'
$destination = (Resolve-Path -LiteralPath $Path).Path
$repo = Split-Path $PSScriptRoot -Parent
$layout = Get-Content (Join-Path $repo 'ui/forms.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$excel = $null; $book = $null; $passed = 0
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.AutomationSecurity = 3
    # Reopen only this synthetic demo, read-only, without enabling file macros.
    $book = $excel.Workbooks.Open($destination, 0, $true)
    $button = $book.Worksheets.Item('AccessTools').Buttons().Item(1)
    $expected = "'" + $book.Name.Replace("'","''") + "'!AtAdminOpen"
    if ($button.OnAction -ne $expected) { throw "Persisted OnAction mismatch: $($button.OnAction)" }; $passed++
    if ($button.Caption -ne ([string]$layout.LaunchCaption)) { throw 'Persisted launch caption mismatch.' }; $passed++
    foreach ($definition in $layout.Forms) {
        $component = $book.VBProject.VBComponents.Item($definition.Name)
        if ($component.Type -ne 3) { throw 'Not a native UserForm.' }; $passed++
        if ($component.CodeModule.CountOfLines -eq 0) { throw 'Persisted form event code missing.' }; $passed++
        if ($component.Designer.Controls.Count -ne $definition.Controls.Count) { throw 'Persisted controls missing.' }; $passed++
        foreach ($control in $definition.Controls) {
            $native = $component.Designer.Controls.Item($control.Name)
            if ($control.Caption -and $native.Caption -ne $control.Caption) { throw "Persisted control caption mismatch: $($control.Name)" }
        }
        $passed++
    }
    $module = $book.VBProject.VBComponents.Item('modAtAdminUi').CodeModule
    if ($module.Lines(1,$module.CountOfLines) -notmatch 'Public Sub AtAdminOpen\(') { throw 'Public launcher missing.' }; $passed++
    if (-not $book.ReadOnly) { throw 'Verification must stay read-only.' }; $passed++
    Write-Output "PASS|DEMO|$passed|0|ManualUI=UNVERIFIED"
    Write-Output ('SHA256=' + (Get-FileHash -LiteralPath $destination -Algorithm SHA256).Hash)
} finally {
    if ($book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
}

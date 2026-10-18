param([Parameter(Mandatory=$true)]$Project, [Parameter(Mandatory=$true)][string]$Root)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
# Designer controls are created during build, so source forms need no binary FRX.
$definitions = @(
    @{Name='frmATResults'; Caption='AccessTools 2.0 — Результаты'; Width=710; Height=455; Controls=@(
        @('Label','lblSummary','',12,12,675,30),
        @('ListBox','lstIssues','',12,46,675,140),
        @('TextBox','txtDetails','',12,196,675,135),
        @('CommandButton','btnAll','Все проблемы',12,344,110,26),
        @('CommandButton','btnCopy','Копировать',130,344,105,26),
        @('CommandButton','btnGo','Перейти',243,344,95,26),
        @('CommandButton','btnAction','Сводка и Superset',346,344,145,26),
        @('CommandButton','btnDiagnostic','Диагностика',499,344,115,26),
        @('CommandButton','btnRepeat','Повторить проверку',12,382,150,26),
        @('CommandButton','btnClear','Очистить отметки',174,382,145,26),
        @('CommandButton','btnSave','Сохранить результаты',331,382,165,26),
        @('CommandButton','btnClose','Закрыть',578,382,109,26)
    )},
    @{Name='frmATTemplates'; Caption='AccessTools 2.0 — Шаблоны и правила'; Width=560; Height=355; Controls=@(
        @('Label','lblSearch','Найти шаблон по названию',12,12,520,20),
        @('TextBox','txtSearch','',12,36,520,26),
        @('ListBox','lstTemplates','',12,74,520,125),
        @('Label','lblHint','Выберите шаблон, затем проверку или его настройки.',12,211,520,24),
        @('CommandButton','btnCheck','Проверить заявку',12,245,160,28),
        @('CommandButton','btnRules','Открыть правила',184,245,160,28),
        @('CommandButton','btnSources','Выбрать источники',356,245,176,28),
        @('CommandButton','btnClose','Закрыть',422,286,110,26)
    )}
)
foreach ($definition in $definitions) {
    $designer = $null
    $component = $Project.VBComponents.Add(3)
    try {
        $component.Name = $definition.Name
        $component.Properties.Item('Caption').Value = $definition.Caption
        $component.Properties.Item('Width').Value = [string]$definition.Width
        $component.Properties.Item('Height').Value = [string]$definition.Height
        $designer = $component.Designer
        foreach ($spec in $definition.Controls) {
            $control = $designer.Controls.Add(('Forms.' + $spec[0] + '.1'), $spec[1], $true)
            try {
                $control.Left=$spec[3]; $control.Top=$spec[4]; $control.Width=$spec[5]; $control.Height=$spec[6]
                # Use the Excel host's standard form font.
                if ($spec[0] -notin @('ListBox','TextBox')) { $control.Caption=$spec[2] }
                if ($spec[1] -eq 'txtDetails') { $control.MultiLine=$true; $control.WordWrap=$true; $control.ScrollBars=2; $control.Locked=$true }
            } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($control) }
        }
        $code = [IO.File]::ReadAllText((Join-Path $Root ('src/ui/forms/' + $definition.Name + '.formcode')))
        $component.CodeModule.AddFromString($code)
        Write-Output ('[FORM] ' + $definition.Name)
    } catch {
        Write-Output $_.ScriptStackTrace
        throw
    } finally {
        if ($designer) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($designer) }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($component)
    }
}

function Import-AtProject($Project, [string]$Repo, [switch]$IncludeTests) {
    $layout = Get-Content -LiteralPath (Join-Path $Repo 'ui/forms.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($form in $layout.Forms) {
        $component = $Project.VBComponents.Add(3)
        $component.Name = $form.Name
        $component.Properties.Item('Caption').Value = $form.Caption
        $component.Properties.Item('Width').Value = [string]$form.Width
        $component.Properties.Item('Height').Value = [string]$form.Height
        $designer = $component.Designer
        $designer.BackColor = 16448250; $designer.Font.Name = 'Segoe UI'; $designer.Font.Size = 10
        foreach ($definition in $form.Controls) {
            $control = $designer.Controls.Add(('Forms.' + $definition.Type + '.1'), $definition.Name, $true)
            foreach ($property in $definition.PSObject.Properties) {
                if ($property.Name -in 'Name','Type') { continue }
                if ($property.Name -eq 'FontSize') { continue }
                $value = $property.Value
                if ($value -is [long]) { $value = [int]$value }
                $control.($property.Name) = $value
            }
            if ($definition.Type -eq 'Label') { $control.WordWrap = $true }
        }
    }
    $files = @(Get-ChildItem (Join-Path $Repo 'src') -File | Where-Object { $_.Extension -in '.bas','.cls','.frm' })
    if ($IncludeTests) { $files += @(Get-ChildItem (Join-Path $Repo 'tests') -Filter '*.bas') }
    foreach ($file in $files) {
        $source = [IO.File]::ReadAllText($file.FullName)
        $name = [regex]::Match($source,'Attribute VB_Name = "([^"]+)"').Groups[1].Value
        if (-not $name) { throw "Missing module name: $file" }
        $source = ($source -split "`r?`n" | Where-Object { $_ -notmatch '^(Attribute |VERSION |BEGIN$|END$|  MultiUse)' }) -join "`r`n"
        if ($file.Extension -eq '.frm') { $component = $Project.VBComponents.Item($name) }
        else { $component = $Project.VBComponents.Add($(if ($file.Extension -eq '.cls') {2} else {1})); $component.Name = $name }
        $component.CodeModule.AddFromString($source)
    }
    return $files + @(Get-Item (Join-Path $Repo 'ui/forms.json'))
}

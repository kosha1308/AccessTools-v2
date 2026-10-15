# AccessTools-v2

Independent VBA implementation of the AccessTools typed rule language.
Specification: AccessTools technical specification v0.13, appendices N-P.

## Status

Stage 1: compiler foundation in development. No production checks are replaced.
AccessTools 0.2.x is a reference line, not a build/runtime dependency.

## Developer checks (Windows)

Run from the repository root in Windows PowerShell 5.1:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tools/StaticValidate.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/tooling/StaticValidate.Tests.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tools/Test.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/tooling/ExcelGate.Tests.ps1
```

ExecutionPolicy applies only to each developer process; these scripts do not change
machine/user execution policy, Excel Trust Center, registry, or security settings.
The Excel checks require installed desktop Excel and existing permission to access
the VBA project object model. Missing prerequisites fail the gate. Static CI does
not replace local Excel verification (hosted GitHub runners do not include Excel).

The runner creates its own hidden Excel process and an empty temporary workbook,
imports only this repository's source/tests, explicitly compiles VBAProject, then
runs the fixed test entry point. A timeout or compile error fails the gate. Cleanup
only terminates the created Excel process after verifying its ID and start time.
Existing user workbooks are not used. Logs, staged imports and JSON reports are
written beneath ignored artifacts/. Source files use UTF-8; VBE import converts to
the local ANSI codepage and rejects unrepresentable characters instead of losing data.

Production Runtime will remain local/offline with no PowerShell, Python, Node,
external DLL or administrator requirement. The scripts above are developer tools.

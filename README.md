# AccessTools-v2

Independent VBA implementation of the AccessTools typed rule language.
Specification: AccessTools technical specification v0.13, appendices N-P.

## Status

Stage 1: compiler profile 1.0.0 implemented and verified in desktop Excel.
Lexer/parser/serializer, catalog/signatures, type/scope checking, source-schema
lookups, compiled named/default bindings and limits are present. The generated
language reference is [catalog.md](docs/dsl/catalog.md). See docs/stage1-status.md
for acceptance evidence, scope and the boundary before Runtime evaluation.
No production checks are replaced.
AccessTools 0.2.x is a reference line, not a build/runtime dependency.

Stage 2A Runtime core is implemented on feat/runtime-core and verified in desktop Excel: 1264 assertions PASS / 0 FAIL, including the compiler baseline. See [runtime-core-status.md](docs/runtime-core-status.md) for scope, review findings and acceptance evidence. Domain/source/row capabilities follow in stage 2B; no production checks are replaced.

## Developer checks (Windows)

Stage 2B.1 adds verified IP/port domains, frozen interval payloads and bounded
normalization/evaluation; see [network-domains-status.md](docs/network-domains-status.md).
The user-facing readiness target includes NETWORK_ACCESS checks, Superset and the
coordinator summary: [first-use milestone](docs/first-use-network-access.md).
Stage 2B.2 adds verified row evaluation and correlated table predicates:
[row-runtime-status.md](docs/row-runtime-status.md). Request adapters and coordinator
delivery remain to implement.

Run from the repository root in Windows PowerShell 5.1:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tools/StaticValidate.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/tooling/StaticValidate.Tests.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tools/Test.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/tooling/ExcelGate.Tests.ps1
```

The Excel runner also compares the generated registry documentation with the
committed snapshot. After intentional registry metadata edits, regenerate it with
tools/Test.ps1 -UpdateCatalog, then run tools/Test.ps1 normally to verify no drift.

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

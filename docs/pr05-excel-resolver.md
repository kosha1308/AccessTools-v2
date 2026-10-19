# PR-05 — универсальный Excel Resolver

База: `ee183bb7efdec14d2b727a2a2629e157a62cb617`. Только чтение явно переданного открытого Workbook. Нет Writer, UI, распознавания шаблона, внешних справочников или табличного исполнения AST. Схемы XML 1.0.0/2.0.0 и их смысл сохранены; Evaluator по-прежнему исполняет только 2.0.0.

## Контракт Resolver → Evaluator

```vb
Dim resolver As CAtExcelResolver, context As CAtEvalContext
Dim diagnostics As Collection, evaluator As New CAtEvaluator
' columnOffsets: Dictionary ColumnId -> 1-based offset within table rectangle.
' For a template without tables pass an empty AtDictionary().
Set resolver = AtNewExcelResolver(package, templateId, columnOffsets)
Set context = resolver.Resolve(workbook, diagnostics)
If context Is Nothing Then
    ' Fatal diagnostic: do not evaluate or infer success.
Else
    Set result = evaluator.Evaluate(package, expressionId, context)
    ' State, diagnostics AND IsComplete matter; Value alone is not PASS.
End If
```

`AtNewExcelResolver` validates and deep-copies the package through the existing XML reader and copies column offsets. Optional trusted `catalog` has the PR-03 signature contract; default is the PR-04 registry catalog. It does not register executable operations. Malformed configuration, unknown template/binding or non-integer offset → host error 5, before Excel access. Changes to the caller's package/bindings do not change an existing Resolver; create a new one after editing locators.

`Resolve(workbook, ByRef diagnostics)` returns only fields owned by the selected template, as `FieldId → CAtEvalResult`, inside the existing immutable context. Ordinary field failures remain typed inputs; the Evaluator preserves diagnostics and creates unresolved EvaluationRequirements when observing non-OK fields. Fatal context construction failure → Nothing plus `R_CONTEXT`/`R_CONTEXT_LIMIT`; no partial context. An empty diagnostic collection from Resolve does **not** mean every field is OK: inspect input/evaluation envelopes.

`ReadField(workbook, fieldId)` returns a typed detached envelope. `ReadTable(workbook, tableId)` returns `CAtTableSnapshot`: TableId, physical StartRow, RowCount, ColumnIds, Status, ReadCell(1-based rowIndex, ColumnId), IsEmptyRow. All result/collection getters return copies; no Workbook/Range references escape. Invalid snapshot row/column indexing is a host error 5.

Provenance records distinguish Input, Locator/version, Workbook name, Requested address, canonical Cell address and FormulaResult. Diagnostics have ObjectId (field/locator or `ColumnId/physicalRow`), code and message; Resolver NodeId is empty because it does not visit AST. TableColumn scalar dependencies retain the logical **FieldId**, with column/row retained in provenance. Snapshots are synchronous captures, not an Excel transaction or an automatic freshness monitor.

## Локаторы и объединения

| Strategy | Implemented contract |
| --- | --- |
| A1 | Validated literal Sheet and bounded rectangle. Scalar may be one cell, or a rectangle entirely inside one MergeArea. Multiple ordinary cells → CONFLICT; never silently select the first |
| DefinedName | Enumerate names of the explicit workbook, accept one workbook name or one name local to declared Sheet. Both scopes present → CONFLICT. Only direct **absolute** A1 rectangle on declared Sheet, validated by AtValidAddress through XFD1048576 (both endpoints and ascending order); invalid bounds → INVALID / R_NAME_REFERENCE before Range access. Missing/broken reference → MISSING. Relative/mixed names, expressions, unions and other unsupported reference forms → INVALID. External workbook syntax → UNAVAILABLE without dereferencing/opening; no Evaluate, RefersToRange or Range(name) |
| LabelOffset | Exact binary match of cached string values in Sheet.UsedRange, including hidden rows/columns. One match required: zero → MISSING, multiple → CONFLICT. Offset is from label MergeArea's top-left. Out-of-grid result → MISSING. Search is bounded before loading values; no Find or changes to Excel Find settings |
| TableColumn | Logical Table/Column refs plus explicit host column-offset binding. Sheet must match table bounds locator, otherwise INVALID. Exactly one declared data row required for scalar read; multiple rows → CONFLICT, with no implicit ANY/ALL/first-row selection. Structural table failures block the field; invalid values in other columns do not contaminate the selected typed cell or its diagnostics. Public table.Status still reports aggregate cell quality |

Scalar reads from any member of a merged area use its upper-left cell, even when outside the requested single-cell address. A subrectangle entirely inside that merge has the same meaning. A rectangle spanning additional cells is CONFLICT. No merge is changed. **Any merged cell inside a table is CONFLICT**, including vertical merges, horizontal merges and empty merged cells; values are not duplicated across logical rows or columns.

Tables use a single bounded data rectangle from A1/DefinedName/LabelOffset. It contains data only: no inferred header, last row, primary column or end marker. Every physical row is retained, including fully empty interior/trailing rows, hidden/filtered rows and partially filled rows. Column order follows declared definitions; physical positions come only from explicit relative offsets. Missing binding/out-of-bounds offset → MISSING; duplicate offsets in one table → CONFLICT. TableColumn cannot recursively define table bounds. Cell errors remain inspectable; structural merges (including unmapped columns) set Status CONFLICT, otherwise Status adopts the first non-OK/non-EMPTY cell state in row/column order and retains bounded diagnostics. Status OK means readable table structure, **not** rule success; EMPTY cells remain EMPTY.

**Current schema gap:** PR-02 Column has DataType but no physical index/header/locator property. PR-05 therefore requires a separate host binding dictionary, not serialized in existing XML. DisplayName is never used as an address/header. Persisted physical column bindings, dynamic boundaries and Excel ListObject semantics need a separately versioned contract in a later authorized task; they are not silently added to 1.0.0/2.0.0. No special handling of SIMPLE/EQUIPMENT_RETURN exists in production.

## Значения, состояния и безопасность

Only `.Value2` is read for data; `.HasFormula` supplies provenance. No Calculate, Evaluate, UpdateLinks, Open, Activate, Select, Save or writes to Excel occur in Resolver. It does not toggle application flags or restore Saved artificially. Formula results are the values currently exposed by Excel: a dirty formula under manual calculation stays cached; freshness is not asserted. Existing Excel/host recalculation or concurrent automation is outside this synchronous read contract.

| Declared type/input | Boundary codec |
| --- | --- |
| Text | Native string, preserving Unicode/whitespace; no automatic numeric formatting/stringification |
| Bool | Native Boolean → canonical `true`/`false`, or exactly canonical text; numeric 0/1 rejected |
| Int64 | Canonical full-range text, or integral native number with absolute value ≤ 9007199254740991; larger numeric values rejected for precision safety |
| Decimal / Date / DateTime | Canonical text accepted by PR-03. Numeric decimal/Excel serial dates are rejected; no locale, date-system, format or timezone guessing |
| Email / Hostname / IpSet / PortSet | No codec yet: nonempty data → INVALID / R_CODEC_UNSUPPORTED, declared type preserved, no Text substitution |

Existing empty cell or zero-length string/formula result → EMPTY, before codec conversion. Excel error values → INVALID / R_EXCEL_ERROR. Unresolvable field/sheet/name/bounds → MISSING; unusable Workbook/COM resource → UNAVAILABLE; ambiguity → CONFLICT. INVALID/UNAVAILABLE do not carry a success payload. ERROR/NOT_APPLICABLE are not invented for ordinary input values. Resource exhaustion → LIMIT, or fatal R_CONTEXT_LIMIT when the immutable context cannot be built.

Hard budgets: package limits from PR-02/03; 100000 worksheet/name/search/cell probes per public read or whole Resolve; at most 2000 names per lookup; table ≤1000 rows, ≤256 physical columns, ≤100000 bounded cells; table value/diagnostic/provenance character sum ≤1048576. Over-budget table returns zero rows, never a partial successful table. Context retains its existing ≤2000 fields, ≤65536 characters/value, ≤128 diagnostic/provenance records/field and ≤1048576 total characters. Table diagnostics are bounded at 128; cell diagnostics remain individually available.

## Проверка и ручная приёмка

`tests/Run-Tests.ps1` imports sources into its own hidden Excel, keeps all 256 PR-02/03/04 regressions and runs `AtRunResolverTests`. Tests create two synthetic workbooks and save `simple-resolver.xlsx` / `equipment-return-resolver.xlsx` only under the ignored per-run artifacts directory. No user requests, real rules or confidential data are used. Test-only writes, comments, merges, Calculate and Save prepare/check fixtures; production Resolver does none of them.

Coverage: scalar/merged/empty values, formulas/errors/manual cached results, missing resources, name scopes/unsafe names, exact labels, invalid/unsupported codecs, bounded empty/partial/multi-column tables and merges, frozen snapshots/configuration, moved field with stable RuleId, two templates through the same Excel → context → Evaluator chain, requirements, resource caps, repeated reads preserving formulas/formatting/comments/structure, Saved True/False, active workbook, application flags and workbook count.

Actual verification: **394 PASS, 0 FAIL = 96 schema +76 AST +84 Evaluator +138 Resolver assertions** in real Excel 16.0 build 17932, Office x64, Windows x64. Evidence: ignored `artifacts/run-104902542b0b4e4d90ec43781d0e9122` (host, source hashes, four result files and detailed transcripts). The initial late-bound Resolver test failed before implementation, while all 256 earlier regressions passed. The 23 review regressions cover independent TableColumn reads versus aggregate table status, retained structural failures, and DefinedName boundaries. On this Excel build Names.Add rejects the four invalid out-of-grid test references at creation; helper rejection and end-to-end resolution of the last valid Excel cell are verified. Office x86 and interactive UI acceptance are not claimed.

Manual acceptance: run the same test runner in a separate supported Excel environment; inspect `host.txt`, four result files and `resolver-assertions.txt`. Open only the generated synthetic books to inspect their fixtures. In a trusted development project, import src modules/classes, create a Resolver with an explicit workbook and inspect returned states before/after editing a field; verify existing annotations and Saved state stay unchanged. This is not an Admin UI or shipping XLAM.

# PR-06 — безопасный ограниченный Excel Writer

База main `2a9fee3affcbfabdfa0ad822be9a245675cfec9e`, ветка `feat/pr-06-excel-writer`. Writer принимает готовые действия и явно переданный Workbook; не знает TemplateId, RuleOutcome, Evaluator и бизнес-правил. XML-схемы PR-02/03 не изменены.

## Контракт и поддержка

```vb
Dim intents As New Collection, intent As New CAtActionIntent
Dim journal As CAtOwnershipJournal, plan As CAtActionPlan, receipts As Collection
intent.ActionInstanceKey = "ANNOTATION_1"
intent.Sheet = "Request": intent.Address = "B8"
intent.Severity = "ERROR": intent.Message = "Synthetic message"
intent.FillColor = RGB(255, 80, 80)
intent.Highlight = True: intent.AddNote = False
intents.Add intent
' Existing dedicated empty directory on a fixed local disk, chosen by the host.
Set journal = AtNewOwnershipJournal(workbook, journalDirectory)
Set plan = AtPlanActions(workbook, intents, "RUN_1", journal)
Set receipts = AtApplyPlan(plan, journal)
' Recheck: same journal, new intents and RunId. Empty intents clears own regions.
' Explicit cleanup: AtClearOwned(workbook, journal, "RUN_CLEAR")
' Interrupted live session: AtRecoverOwned(workbook, journal, "RUN_RECOVERY")
```

ActionIntent contains stable caller instance key, Sheet/Address, severity, message, RGB color and independent Highlight/AddNote flags. RunId and caller keys use existing uppercase ID syntax. Full instance identity is RunId/caller key/canonical target; retry of the same plan does not duplicate notes. ActionPlan copies scalar records and pins the Workbook and journal session; getters return copies. ActionReceipt contains instance identities, RunId, target, terminal status/code/message and verified actual snapshot. Several intents on one physical area produce a grouped receipt retaining all instance keys. Failed plans have a plan-level receipt; interrupted runs retain earlier receipts and mark unattempted targets BLOCKED_RECOVERY.

Planner validates the whole input before any Excel mutation, normalizes cells/subrectangles within one MergeArea, then groups by worksheet CodeName + canonical address. Messages remain in input order, including identical texts from distinct intents. Highest severity chooses fill: ERROR > WARNING > INFO; equal-priority tie goes to the first highlight intent. Notes use LF line endings; unsupported control characters and oversized payloads are refused. Ordinary multi-cell rectangles outside one MergeArea are not supported; use individual intents.

| Operation | Supported boundary |
| --- | --- |
| Fill/restore | No fill or uniform direct RGB solid fill; no tint/theme/pattern/conditional formatting. Per-member merged-cell fill is inspected; nonuniform MergeArea is refused |
| New Legacy Note / replace own note / remove own note | Only the proven Windows Excel 16.0 build 17932 **Office x64**, BIFF8 `.xls` profile. No existing user note is edited. New note capture includes text, author, geometry, frame, font/color/palette/system linkage, fill/line, supported shadow metadata and Z-order; mixed fonts, shadow/3D or unsupported fill profiles are refused |
| Multiple messages | One joined note per canonical MergeArea; no deduplication by message text |
| Modern `.xlsx`, new unsaved workbook | Fill-only works. New notes are CAPABILITY_UNSUPPORTED in this PR: installed Excel's threaded API returns 1004 (not supported). Writer never converts/saves a workbook to enable notes |
| Other Office versions/builds, Office x86 | New notes are gated before mutation until a capability profile is proved. Fill-only still uses runtime capture/refusal; x86 execution not tested |

At most one owned note area per worksheet, including old regions awaiting cleanup: removing/recreating another note can change Z-order and invalidate its fingerprint. Moving a note to another area requires explicit clear then a new plan. Multiple fill-only areas and notes on different worksheets are independent. This is a deliberate capability limit, not template-specific logic.

User Notes are refused even when their text contains “AccessTools” and their fill equals the requested color. Native threaded comments, protection of contents/drawings/workbook structure, CF (including its priority), themes, patterned/tinted fills, unsupported note profiles and merged nonuniform fills are never flattened or converted. If an existing owned target no longer matches, outcome is CONFLICT and user changes remain. Even when threaded inspection is unavailable, fill-only never writes/deletes any comment; note operations remain format/profile-gated.

## Сохранность и владение

Writer writes only Interior properties in the supported fill profile and its own Legacy Note objects. It never writes Value2/formulas, changes merges/sheets, opens another workbook, activates a sheet or saves the workbook. Application flags are not toggled. Excel naturally marks a changed workbook dirty; Writer never forges Saved=True. Font, borders, number format, data, formulas and user comments are not restoration targets and are left intact.

Before planning/applying/clearing: capture current per-cell values, formulas, merge geometry and touched properties. Before every physical operation: persist PENDING with original baseline, StepBefore, known expected-next state and intent. After: read Excel again, compare expected versus actual, persist acknowledgement. Note defaults cannot be predicted before creation; ADD_NOTE has no authoritative expected-next shape until full capture/acknowledgement. An interrupted unacknowledged note is **not** adopted by matching its text.

Cleanup/replacement requires current full supported fingerprint to match journal Expected. Original baseline survives replacement/new runs. Recheck plans reconcile obsolete regions from the journal's original targets, not current locators. The whole plan is revalidated immediately before the first mutation. A new external change during apply or an unverified mutation stops dependent work and marks RECOVERY_REQUIRED; earlier successful receipts do not mean complete output. No workbook-wide rollback/ACID guarantee is claimed.

## Физический журнал

Dedicated caller-selected directory on a fixed local disk; UNC/mapped/removable drives are refused through FileSystemObject.DriveType. No hidden worksheet, Names, CustomXML or workbook metadata is added. `session.atj` is an exclusive append-only session lease. Sequential `ATJ_<session>_<revision>.atj` records are UTF-16 XML, format 1, containing scalar key/value lines with four-hex-digit UTF-16 units. Nested snapshots use the same scalar encoding: Unicode and CR/LF survive exactly. Encoding is **not encryption**; local files contain the captured data needed for guards/recovery and must remain in an appropriately protected location, never in Git.

Each record carries BookSession, informational BookName/BookPath/BookFormat, RunId, action keys, target, original Before, Observed, StepBefore, Next, Expected, Actual, phase and lifecycle state. Exact live Workbook object identity is authoritative; path/name alone never grants ownership. Unsaved workbooks use that live session identity; user SaveAs is supported for fill-only when the captured format stays compatible. Closing/reopening does not silently rebind a journal.

Before each mutation, a new file is created without overwrite, closed, read back and compared. Every verification checks both known file contents and directory inventory; unknown/missing/modified files block writing. Failure to persist is not Excel success. This proves tested ordered file writes/readback, **not fsync, power-loss durability, cryptographic integrity or a transaction across NTFS and Excel**.

Budgets: ≤32 intents/plan, ≤32 distinct journal targets, ≤64 cells/region, ≤512 note/message characters, ≤64K packed capture (≤32K before grouped note actions). Normal journal ≤256 revisions (host may lower this), ≤8M UTF-16 characters total; individual record ≤2M characters. Recovery has a separate reserved ceiling of 1024 revisions/64M characters so normal exhaustion does not itself prevent rollback. Physical disk failure, exhausted recovery reserve or corrupted evidence still requires manual intervention; no silent truncation/rotation/deletion.

## Восстановление

1. Any interrupted/unverified write sets RECOVERY_REQUIRED and blocks new plans. Retain the journal and owning live Workbook; do not delete `.atj` files to bypass the block.
2. Call AtRecoverOwned with that same Workbook/journal. It checks all records and all targets first. Only a fully journaled Expected or known Next snapshot is eligible; acknowledged partial fills can be restored. Recovery steps are themselves journaled and verified.
3. User edits, an unknown post-ADD_NOTE state, lost/changed files or identity mismatch → CONFLICT/RECOVERY_REQUIRED, no destructive fallback. Inspect/export a copy manually using the recorded target and before/after evidence.
4. After process death/loss of the in-memory journal, new journal creation sees the surviving `.atj` files and blocks. Automatic import, workbook rebinding and automatic recovery after Excel/process/power crash are **not supported**. Do not assert ownership from the serialized text/color/path. No automatic startup scan or UI is provided.

## Проверка и ручная приёмка

`tests/Run-Tests.ps1` keeps 394 PR-02–05 regressions and runs AtRunWriterTests in its own Excel on synthetic temporary books only. The first Writer bootstrap was RED before implementation. Tests cover fill/no-fill and custom RGB restore, preserved data/formulas/font/borders, merge grouping/priority/all messages, new/updated/removed own notes, no duplicates, user-note protection, CF priorities, sheet protection, nonuniform merges, user edits/shape/font conflicts, snapshots, unsaved workbook/SaveAs, unchanged application flags, stale plans, interrupted steps/multiple targets, journal collision/tampering/orphans, normal-budget exhaustion and reserved recovery.

Actual verification: **458 PASS, 0 FAIL = 96 schema +76 AST +84 Evaluator +138 Resolver +64 Writer**, real Excel 16.0 build 17932, Office x64 / Windows x64. Evidence: ignored artifacts/run-f01f7843ebf748eaa589b539c9b03861 (host, source hashes, five result files, writer-assertions.txt). Native threaded comment creation and the Legacy Note font ThemeColor setter are unavailable in the installed Excel; these two positive cases are UNVERIFIED, not PASS. Same-RGB palette/system linkage changes are tested. Unsupported-API note refusal is tested. Office x86, interactive UI acceptance, actual process termination/power loss, cross-process recovery, network-drive creation and arbitrary Excel shapes/styles are not claimed. A later capability extension needs its own proof; PR-07 is not started.

Manual acceptance: run the same test runner on a separate supported Excel host; inspect `host.txt`, five result files and detailed writer assertions. On synthetic data, apply twice, modify a marked cell/note, verify CONFLICT without deletion; undo the user change and clear. Run with mutation budget 1, verify RECOVERY_REQUIRED and explicit guarded recovery. Check CF/user notes/protected targets are refused before any partial formatting. Manual acceptance remains separate from automated PASS.

API reference: [Microsoft Font.ThemeColor](https://learn.microsoft.com/en-us/office/vba/api/excel.font.themecolor) documents an invalid-request runtime error for an unthemed font. The tested profile records this specific absence as NO_THEME; unrelated query errors refuse capture. System scheme identity (80/81), direct RGB type, palette index and tint are compared, not just displayed RGB.

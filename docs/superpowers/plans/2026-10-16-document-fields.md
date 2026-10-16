# Document Fields Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL superpowers:executing-plans. Native execution; delegated technical decisions.

**Goal:** Read described document fields with merge-aware offsets and source-cell trace.
**Architecture:** Immutable field spec/result, bounded locator and additive context metadata.
**Tech Stack:** Offline VBA, existing real Excel test harness.
**Spec:** docs/superpowers/specs/2026-10-16-document-fields-design.md.

## Global Constraints

Baseline9733/0. No input mutation. All raw fields TEXT. Search<=1000000 cells.
No release claim; first-use still requires full NETWORK_ACCESS+Superset+coordinator.

## Review Focus

Signed offsets at merged boundaries; duplicate normalized labels; blank vs missing vs CVErr;
source references after closing workbook; document-field trace inside active row context.

### Task1 — immutable single-field locator

Files: src/excel/ATExcelFieldSpec.cls, ATExcelFieldValue.cls, ATExcelFieldReader.cls;
modATExcelText.bas (shared cell-reference helper), ATExcelTableReader.cls (use helper);
tests/vba/modATExcelFieldTests.bas; modATTestMain.bas.
Interfaces: spec.Initialize(fieldId,anchorText,firstRow,lastRow,firstColumn,lastColumn,
rowOffset,columnOffset); read-only properties with matching names.
reader.Read(ws As Worksheet,spec As ATExcelFieldSpec) As ATExcelFieldValue;
value.FieldId/Found/SheetName/AnchorRef/CellRef/Result (immutable; Result cloned).
AT_ExcelCellReference(sheetName,row,column) As String, replacing table reader private helper.
- [ ] RED real-sheet tests: four directions, diagonal/exact offsets, merged value/anchor,
  zero target/row/column bounds, duplicate/missing normalized anchor, unsafe specs,
  empty/blank formula/CVErr/zero/raw malformed text, reuse and source preserved.
- [ ] Implement bounded array locator and immutable snapshots; full Excel GREEN, commit.

### Task2 — document metadata through runtime

Files: src/runtime/ATEvaluationContext.cls, ATEvaluator.cls;
tests/vba/modATDocumentFieldTests.bas, test main.
Interfaces: context.AddField(id,result,Optional cellRef=""); CellRef(id) As String;
Friend RetainField(id,result,Optional cellRef="") preserves snapshot. LogNode document
FIELD metadata resolves through mContext, without reading skipped values.
- [ ] RED: reader result -> context -> compiled DSL; real CellRef in success and error;
  metadata preserved across context Snapshot; row/document address separation.
- [ ] Implement additive metadata, full GREEN; commit. Fresh whole-branch review, fix Important
  with RED/GREEN; static7/7, gate9/9, clean-checkout compile/all tests.
- [ ] Russian status/milestone docs; push new branch and draft PR against feat/network-extraction,
  attach and check latest-head CI. No merge or ready-to-use claim.

Pre-flight: task1 Result feeds task2 AddField with CellRef; signatures consistent.
Ruling: retain dedicated project clone and new feature branch, as on previous stages.

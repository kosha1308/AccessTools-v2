# Excel Table Extraction Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL superpowers:executing-plans. Native execution; user delegated technical decisions.

**Goal:** Read bounded Excel tables into immutable raw TEXT rows with source-cell metadata.
**Architecture:** External column descriptions, unique header locator, one Value2 block and local merge cache; runtime unchanged.
**Tech Stack:** Offline VBA and existing isolated Excel test runner.
**Spec:** docs/superpowers/specs/2026-10-16-network-extraction-design.md.

## Global Constraints

No user workbook mutations, no legacy source/data copies, no runtime dependencies.
Baseline9679/0. Maximum10000 physical rows,1000000 cells,10000 Find candidates/column.
All extracted fields TEXT; no implicit IP/group conversion. Full first-use requires Superset/coordinator.

## Review Focus

Merged header with bottom offset; DIRECT auxiliary vs inherited blank; CVErr row participation;
stale layouts across sheets; formatting/formula-empty at sheet end. Tests below cover all five.

### Task1 — configurable header locator

Files: src/excel/ATExcelTableSpec.cls, ATExcelTableLayout.cls, modATExcelText.bas;
tests/vba/modATExcelLocatorTests.bas; modATTestMain.bas.
Interfaces: spec.Initialize(tableId,firstRow,lastRow,firstColumn,lastColumn);
AddColumn(fieldId,headerText,usage); Locate(ws As Worksheet) As ATExcelTableLayout.
layout.TableId/HeaderFirstRow/HeaderLastRow/Count/FirstColumn/LastColumn;
FieldId(index),ColumnNumber(index),Usage(index),ValidateSheet(ws).
AT_ExcelBlank(raw As Variant) As Boolean; AT_ExcelText(raw) As String; AT_ExcelHeader(text) As String.
- [ ] RED real sheet: reordered and merged headers, unique full header, missing/duplicate/ambiguous,
  invalid profiles, bounds, wrong sheet, immutable layout after spec edits.
- [ ] Implement config validation and bounded array search; header merge anchors/range validation.
- [ ] Full Excel GREEN; commit.

### Task2 — bounded row snapshots

Files: src/excel/ATExcelTableReader.cls; tests/vba/modATExcelReaderTests.bas; test main.
Interfaces: reader.Read(ws,layout,lastDataRow) As ATEvalResult;
reader.LastValueRow(ws,layout) As Long. Reads start at layout.HeaderLastRow+1;
empty allowed lastDataRow=HeaderLastRow. Failure codes EXCEL_*; no partial rows on structural failure.
- [ ] RED synthetic NETWORK_ACCESS-shaped sheets: merged reason skipsblank; key inheritance;
  lower DIRECT with merged aux skipsblank but separate aux retained; blank formula/unicode;
  CVErr gives field ERROR; raw malformed/multiline intact; merged anchor refs and physical ids.
- [ ] RED bounds/cross-table merge/wrong sheet/reuse/format-only last row/empty formula last row;
  input formula/style/merge unchanged, immutable result, compiled COUNT/ALLROWS over extracted rows.
- [ ] Implement single Value2 block, local per-column merge cache, bounded last-value search.
- [ ] Full GREEN; independent whole-branch reviewer required by executing-plans.
  Fix important findings with reproducing RED/GREEN; static7/7, gate9/9, clean-checkout compile/tests.
- [ ] Russian status and first-use checklist; commit; push new feature branch and draft PR
  base feat/runtime-rows, attach, confirm latest-head CI. No merge or release claim.

Pre-flight: locator yields immutable, sheet-bound layout used by reader. No shared runtime changes.
Ruling: retain dedicated clean project clone and new branch, consistent with previous stages.

# Row Runtime Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL superpowers:executing-plans. Native execution in this chat; technical decisions delegated by user.

**Goal:** Execute typed row predicates and correlated table comparisons.
**Architecture:** Immutable rows/collections, compiled bindings, scoped traversal and bounded trace.
**Tech Stack:** Offline VBA, existing Excel runner.
**Spec:** docs/superpowers/specs/2026-10-16-row-runtime-design.md.

## Global Constraints

Preserve9563 baseline, no real workbook I/O, no arbitrary code, no production switch.
Existing limits: steps200000/depth64/list10000/trace20000/text65536. Branch feat/runtime-rows.
User-ready target remains full NETWORK_ACCESS + Superset + coordinator output.

## Review Focus

Unqualified field exact alias (task1); caller mutations (task2); DOCUMENT alias shadow
(task3); non-OK plus trace limit (task3); unused bad row fields with COUNT/short circuit(task3).

### Task1 — compiled binding

Files: ATAstNode.cls, ATCompiler.cls, tests/vba/modATRowBindingTests.bas, test main.
Interfaces: AST FieldId/TableId/RowAlias/NamedScope Strings. FIELD binds from resolver
result and selected lexical alias. NamedScope copied from definition; no serialization change.
- [ ] Write binding tests: explicit/unqualified/document, named DOCUMENT/ROW and nested scopes.
- [ ] Run RED, implement metadata, run full GREEN and commit.

### Task2 — row values

Files: new ATRuntimeRow.cls; ATRuntimeValue.cls; tests/vba/modATRowModelTests.bas.
Interfaces: row.Initialize(tableId As String,rowId As String); AddField(id As String,
result As ATEvalResult,Optional cellRef As String=""); Snapshot() As ATRuntimeRow;
ResolveField(id As String,Optional budget As ATEvalBudget=Nothing) As ATEvalResult;
CellRef(id As String) As String; read-only TableId/RowId.
Value.InitializeRows(tableId As String,rows As Collection); Row(index As Long) As
ATRuntimeRow; RowTableId As String; Count handles rows/list; Clone shares frozen rows only.
- [ ] RED tests: immutable input/snapshot/getters, table mismatch, duplicate RowId, guarded payload.
- [ ] Implement row snapshots and budget validation before cloning; full GREEN and commit.

### Task3 — traversal and trace

Files: ATEvaluator.cls, modATRuntimeDispatch.bas, ATTraceEntry.cls; new
tests/vba/modATRowEvalTests.bas; update old unsupported row fixtures to remaining unsupported signatures.
Interfaces: evaluator scoped aliases plus current-row identity, EvaluateRows(node,frame),
EvaluateNamed(node,frame), ResolveBoundField(node); trace optional RowId/TableId/FieldId/CellRef properties.
Read bound fields by FieldId/RowAlias; missing alias/field/type gives stable RUNTIME_ROW_* diagnostics.
- [ ] RED end-to-end fixtures ALLROWS/ANYROW/NOROWS/FILTER/COUNT/EXISTS; empty truth table;
  correlated IP collections, unqualified field, nested named scopes/defaults and caller shadow.
- [ ] Add lazy/non-OK matrix, missing/type/invalid fields, input/result immutability, depth/row/steps/trace
  limits, reused evaluator after errors and skips. Per-row NOT_EVALUATED metadata has no payload.
- [ ] Implement bounded traversal/restoration and collection handling. Unsupported sources/regex,
  DISTINCT_COUNT and collection equality still fail before reading operands. Full GREEN.
- [ ] Independent review; fix with reproducing tests. Static7/7, gate9/9, clean clone compile/full tests.
  Commit and draft PR against runtime-domains; attach and verify latest-head CI.

### Task4 — user milestone

- [ ] Russian status with exact evidence; update first-use checklist. Rows are values in memory;
  reading Excel form/merged cells, Superset/coordinator and packaged UI remain required.

Preflight: task1 metadata consumed by3, task2 frozen collection consumed by3; no dependency cycle.
No new approval roundtrip: user explicitly delegates core technical decisions and asks to continue.

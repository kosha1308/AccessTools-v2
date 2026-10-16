# Superset Enrichment Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL superpowers:executing-plans. Native implementation; delegated technical decisions.

**Goal:** Apply revised TZ to visible linked Superset data including SKUF/support team.
**Architecture:** External layout + existing immutable raw table reader + bounded text sets/buckets.
**Tech Stack:** Offline VBA, real Excel test harness.
**Spec:** docs/superpowers/specs/2026-10-16-superset-enrichment-design.md.

## Global Constraints

Baseline9796/0. No source mutation or private legacy code/data copied.
10000 source rows/summary ids/set entries,1000000 cells,200000 link/assignment steps,
32767 joined text chars. No IP matcher/refinement/writer or release-ready claim.

## Review Focus

Exact1 vs10 links; duplicated source links don't inflate count; hidden field errors ignored;
multiple hostname ambiguity; Cyrillic/whitespace headers and value dedupe; no partial output on limits.

### Task1 — text sets and per-number bucket

Files: src/superset/ATSupersetTextSet.cls, ATSupersetBucket.cls;
tests/vba/modATSupersetSetTests.bas, test main.
Interfaces: set.AddLines(text); Text(Optional uniqueOnly=False) As String; Count As Long.
bucket.Initialize(summaryId); AddMatch(values As Variant of5 strings);
Snapshot() As ATRuntimeRow with MATCH_COUNT and5 named fields.
- [ ] RED multiline/case/blank/first-order values; unique hostname vs all values; snapshot immutable;
  text/set limits and zero-match row. Implement minimal classes; full Excel GREEN; commit.

### Task2 — visible linked source adapter

Files: src/superset/ATSupersetEnricher.cls; tests/vba/modATSupersetReadTests.bas; test main.
Interface: Read(ws As Worksheet,layout As ATExcelTableLayout,lastDataRow As Long,
summaryIds As Collection) As ATEvalResult. Six semantic fields as spec, id/usage schema validation.
Uses ATExcelTableReader.Read for bounded bulk values and Rows.Hidden for visibility metadata.
- [ ] RED real Excel table with extra-space Russian headers, sixcolumns, IDs1/2/3/10,
  filtered and manuallyhidden rows, multiline/new fields, duplicate links, missing summary id,
  hidden vs visible CVErr, malformed links, exact matching, no source mutation, rerun, closed-source result.
- [ ] Implement local state only, full GREEN; fresh whole-branch review; fix Important RED/GREEN.
  Static7/7, gate9/9 and clean-checkout Excel acceptance.
- [ ] Update Russian first-use checklist with new required fields and smoke gate; push draft PR
  against feat/document-fields, attach and verify latest-head CI. No merge.

Pre-flight: task1 bounded bucket snapshot consumed by task2; no other runtime interface change.
Ruling: first implement revised enrichment contract; pending template fingerprint remains scheduled.

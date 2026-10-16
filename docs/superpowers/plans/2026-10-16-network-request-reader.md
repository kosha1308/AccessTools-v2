# Network Request Reader Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL superpowers:executing-plans. Native implementation, delegated technical choices.

**Goal:** One verified raw request snapshot for NETWORK_ACCESS v3.0.
**Architecture:** Shared form metadata, existing bounded readers, guarded immutable context/result.
**Tech Stack:** Offline VBA and real Excel tests.
**Spec:** docs/superpowers/specs/2026-10-16-network-request-reader-design.md.

## Global Constraints

Baseline9889/0. Updated TZ SKUF/support requirements preserved. No source/settings mutation.
Existing profile1million scanned cells,10000-row/1million-cell reader limits remain.
Only raw data here; no implicit date policy, full business checks or release-ready claim.

## Review Focus

DB account C10; numeric vs text date and workbook Date1904; lower header shifts;
merged auxiliary blanks vs actual rows; rejection cannot expose partially built context.

### Task1 — preserved native value and shared metadata

Files: ATExcelFieldValue.cls, ATExcelFieldReader.cls; modATNetworkProfile.bas;
tests/vba/modATExcelFieldTests.bas and new modATNetworkReadTests.bas.
Interfaces: FieldValue.RawValue2 As Variant scalar, optional rawValue2 in Friend Initialize;
AT_NetworkDocumentFieldsV3() As Collection; AT_NetworkTechnicalSpecV3/AddressSpecV3()
As ATExcelTableSpec with RAW field/table IDs fixed by spec. Profile and readers share header arrays.
- [x] RED native scalar CVErr/numeric/closure immutability and metadata factory tests.
- [x] Implement scalar retention and metadata factories, full Excel GREEN, commit.

### Task2 — verified extraction and runtime context

Files: src/templates/ATNetworkReadResult.cls, ATNetworkRequestReader.cls;
tests/vba/modATNetworkReadTests.bas; main.
Interfaces: Reader.Read(ws,runId,runDate,timezoneId) As ATNetworkReadResult;
result.State/Diagnostics/Context, initialized once, Context only when OK and immutable.
- [x] RED synthetic full form: all11fields, doc/table cell refs, fresh upper/lower bounds,
  merged spacers, CVErr, EMPTY vs serial/date1904, no partial context after structural/bounds error,
  context validation, rerun, source unchanged/closed and compiled DSL over extracted tables.
- [x] Implement Verify-first raw context, full GREEN; independent review, fix Important RED/GREEN.
  Static7/7, gate9/9, actual original read-only acceptance with unchanged hash, clean checkout.
- [x] Russian status/checklist; new feature branch draft PR against feat/template-gate,
  attach and confirm latest-head CI. No merge/release.

Pre-flight: Task1 metadata/native type consumed by Task2; no generic runtime behavior changes.
Ruling: existing dedicated clone + feature branch retained; no new technical approval loops.

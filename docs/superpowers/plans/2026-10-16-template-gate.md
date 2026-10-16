# Template Gate Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL superpowers:executing-plans. Native execution, delegated technical decisions.

**Goal:** Fail-closed structural template verification before table extraction.
**Architecture:** Sealed declarative anchors/relations, existing header locator, immutable gate result,
guarded table-read entry. External production reference required, no invented NETWORK_ACCESS profile.
**Tech Stack:** Offline VBA and real Excel tests.
**Spec:** docs/superpowers/specs/2026-10-16-template-gate-design.md.

## Global Constraints

Baseline9845/0; updated TZ including SKUF/support remains binding. No source/settings mutation.
256 anchors/1024 relations/1000000 total scanned cells; minimum2 independent anchors/1relation.
No release claim. User asked for location of actual XLSX; generic mechanism independent of reply.

## Review Focus

Dynamic lower table rows; formulas impersonating headers; duplicate facts under different ids;
unsealed/invalid profile must fail closed; gate precedes structural/data reader errors.

### Task1 — profile and immutable decision

Files: src/templates/ATTemplateProfile.cls, ATTemplateGateResult.cls;
tests/vba/modATTemplateGateTests.bas and test main.
Interfaces: Initialize(templateId,profileVersion,sheetName,Optional requireSheetProtection=False);
AddAnchor(id,label,firstRow,lastRow,firstCol,lastCol,mergeRows,mergeCols,Optional requireLocked=True);
AddRelation(fromId,toId,minRows,maxRows,minCols,maxCols); Seal(); Verify(ws) As ATTemplateGateResult.
Result readonly Allowed/Code/Detail/Message, initialized once. Metadata arrays copied/owned.
- [ ] RED real sheets: reference passes; user values/style ignored; changed label/version/merge,
  forbidden position/formula/duplicate/protection rejected; allowed row growth; reuse and no writes.
- [ ] RED config/shape/total-size/id limits, two ids same fact, missing relation ends, sealed mutations.
- [ ] Implement bounded profile checks using ATExcelTableSpec; full Excel GREEN, commit.

### Task2 — guarded reader

Files: src/templates/ATTemplateTableGate.cls; same tests.
Interface: Read(profile As ATTemplateProfile,ws As Worksheet,layout As ATExcelTableLayout,
lastDataRow As Long) As ATEvalResult. Verify first, reject with single localized diagnostic;
only allowed profile invokes existing reader.
- [ ] RED mismatch vs invalid data bounds; valid template reads rows; missing profile fails closed;
  source edits require fresh Verify, no partial data on rejection.
- [ ] Implement guard, full GREEN; fresh independent whole-branch review, fix Important RED/GREEN.
  Static7/7, gate9/9 and clean-checkout compile/tests.
- [ ] Russian status/first-use checklist, preserve missing production reference requirement;
  push draft PR against feat/superset-enrichment, attach, latest-head CI. No merge.

Pre-flight: Task1 decision consumed before Task2 reader; no core/schema changes required.
Ruling: retain dedicated clone and new feature branch; no extra technical approval loop.

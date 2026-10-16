# Network tables/report Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL superpowers:executing-plans; inline, one final fresh reviewer.

**Goal:** Table input checks and a separate Russian Excel report.
**Architecture:** Existing snapshots/parsers, explicit protocol catalog, report aggregation, new-workbook-only writer.
**Tech Stack:** Offline VBA and supervised real Excel; PDF render for visual QA.
**Spec:** docs/superpowers/specs/2026-10-16-network-table-report-design.md.

## Global Constraints

Basef5e3204, baseline10381/0. No source/settings mutation; writer restores EnableEvents.
No protocol defaults or false PASS for groups/missing source. Partial scope always visible.
Limits10000rows/200000steps/65536text,990tableissues,1000protocol literals/128chars.

## Review Focus

Pure IP parser cannot classify textual groups: manual SKIPPED, no false prohibition/pass.
Port0 legal in generic math but not this business check.
Empty/moved headers: current snapshot refs, not hardcoded body rows.
Untrusted formula/ref text: plain text, no arbitrary hyperlinks.
Source/app state after writer failure/repeat and input closure: preserve source/events and no overwrite.

### Task1: table checks and local aggregate

Files reader/normalizer refs; src/checks/ATNetworkTableChecks.cls,
ATNetworkLocalChecks.cls; tests/vba/modATTableCheckTests.bas, main.
- [x] RED rows, bounds, refs, errors/warnings/groups/protocol/catalog/duplicate/limits and6+3 aggregation.
- [x] Implement; real Excel full GREEN; commit.

### Task2: Russian separate-workbook report

Files src/ui/ATExcelReportWriter.cls, modATReportLabels.bas;
tests/vba/modATReportWriterTests.bas, main.
- [x] RED labels/counts/no-source-write/formula safety/hyperlink trust/repeat/events.
- [x] Implement; full GREEN, visual PDF render/inspect sample; commit.
- [x] Fresh review/fix Important RED→GREEN; static7/7/gate9/9, original ReadOnly/hash, clean checkout.
- [x] Status/readiness; push draft PR base feat/network-document-checks, attach, latest-head CI; no merge/release.

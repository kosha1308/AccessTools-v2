# Published source reader Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL superpowers:executing-plans. Native inline implementation and one fresh final reviewer.

**Goal:** Read the supplied Superset/Rulebook schemas as immutable raw snapshots.
**Architecture:** Metadata factory + bounded flat-table reader; reuse existing locator/runtime values. Preserve completed checks and enrichment.
**Tech Stack:** Offline VBA, real Excel integration tests.
**Spec:** docs/superpowers/specs/2026-10-16-network-source-reader-design.md.

## Global Constraints

Base86e2f29, suite10488/0. Header window1–20/1–64;100000 body rows;
1000000 array cells. No source mutation, startup load, invented policy or real data.

## Review Focus

Numeric unnamed index alone is not a Superset record.
Actual Superset exceeds10000 rows: don't reuse request-reader row cap.
Reordered/hidden columns, extra whitespace and repeated headers fail or map predictably.
CVErr stays an error; empty rule tables cannot become effective default policy.
Closed/missing sheets, merged records and out-of-budget data fail without partial rows.

### Task1 — source metadata and raw reader

Create src/sources/modATNetworkSourceSchema.bas, ATNetworkSourceReader.cls,
tests/vba/modATNetworkSourceReadTests.bas; modify test main only.
Interface AT_NetworkSourceSpec(sourceId) As ATExcelTableSpec;
ATNetworkSourceReader.Read(ws,sourceId) As ATEvalResult, table sourceId & "_RAW".
- [ ] RED: eight schemas, reordered whitespace headers, blank index-only rows,
 10001 records, hidden rows, errors, fresh reads/closure, missing/duplicate/unknown,
 merged/bounds failure; observe compile/test failure.
- [ ] Implement minimal schema/reader; full real Excel GREEN; commit.
- [ ] Fresh independent review and reproduce/fix findings; static7/7 gate9/9.
- [ ] Read actual two originals with hashes, clean-copy full standard60s suite.
- [ ] Status/readiness, draft PR against feat/network-address-match, attach, CI.

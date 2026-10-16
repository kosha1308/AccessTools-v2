# Document checks Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL superpowers:executing-plans; inline, one fresh final reviewer.

**Goal:** Six local NETWORK_DOCUMENT checks and immutable Russian issue report.
**Architecture:** Explicit DMY_TEXT date policy, generic immutable report values, template-specific checker over existing reader/normalizer.
**Tech Stack:** Pure offline VBA, supervised real Excel.
**Spec:** docs/superpowers/specs/2026-10-16-network-document-checks-design.md.

## Global Constraints

Basea462767, baseline10265/0. No source/settings mutation, current clock or locale parsing.
Source semantic reference only; no private code/data copied. Full release/Superset remain pending.
Report scope NETWORK_DOCUMENT, six checks only, max1000issues/checks. Date text max65536.

## Review Focus

Date words/multiple dates and adjacent digits: retain compatible explicit parsing, no partial numeric salvage.
Revoke ignores bad validity but system errors still block; unknown request does not imply grant.
Empty/CVErr email: one issue, no duplicate syntax issue.
Malformed snapshot/ref/types: no partial positive outcomes.
Frozen report collections and data values: source edits/closure cannot alter counts/addresses.

### Task1: DMY_TEXT compatibility

Files src/excel/ATExcelDateAdapter.cls; tests/vba/modATExcelDateTests.bas.
Interface existing Normalize gains explicit DMY_TEXT format, otherwise unchanged.
- [ ] RED dates in words/multiple candidates/digit boundaries/invalids/ISO/length.
- [ ] Implement bounded scanner; full Excel GREEN; commit.

### Task2: checks and report

Files src/checks/ATCheckIssue.cls, ATCheckOutcome.cls, ATCheckReport.cls,
ATNetworkDocumentChecks.cls, modATEmailSyntax.bas; tests/vba/modATDocumentCheckTests.bas, main.
Interface Run(source As ATNetworkReadResult) As ATCheckReport; immutable properties per spec.
- [ ] RED full reader pipeline blank/valid/multiple issues, email corpus, request/date/revoke/unknown,
  cell errors, source rejection/malformed schema, report freezing/counts and stable refs after close.
- [ ] Implement six checks and generic immutable values; full GREEN, commit.
- [ ] Fresh review, Important RED/GREEN fixes; static7/7 gate9/9, real reference ReadOnly/hash, clean checkout.
- [ ] Russian status/readiness, push draft PR base feat/network-normalization, attach, latest-head CI. No merge/release.

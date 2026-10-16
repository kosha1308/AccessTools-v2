# Network normalization Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL superpowers:executing-plans. Native inline execution, one final fresh review.

**Goal:** Typed request data for subsequent business checks, preserving raw inputs and refs.
**Architecture:** Independent Excel date adapter, template-specific context transformation, existing domain parsers.
**Tech Stack:** Offline VBA, supervised real Excel.
**Spec:** docs/superpowers/specs/2026-10-16-network-normalization-design.md.

## Global Constraints

No source/settings writes. No locale/current-clock inference. Explicit date format and keyword.
Base7ccdd27, baseline9941/0. Max10000rows and existing runtime value budget. No new business policy.
Updated TZ SKUF/support fields remain required for eventual full release.

## Review Focus

Serial60 and leap day/epoch boundaries: reject false dates and verify actual calendar days.
Digits stored as text: invalid date, no numeric coercion.
Raw error/invalid token: preserve failure, never salvage usable subset.
Re-running normalizer/source closure: same detached immutable snapshot and original refs.
Malformed hand-built contexts/config: stable fail-closed ERROR without partial Context.

### Task1 — explicit date adapter

Files src/excel/ATExcelDateAdapter.cls, tests/vba/modATExcelDateTests.bas, main.
Interface Normalize(raw As Variant,date1904 As Boolean,textFormat As String) As ATEvalResult.
- [ ] RED tests: epochs/bounds/serial60/fractions/leaps/text formats/blank/null/errors/objects/policy.
- [ ] Implement, full Excel GREEN, record and commit.

### Task2 — typed NETWORK_ACCESS context

Files src/templates/ATNetworkRequestNormalizer.cls, tests/vba/modATNetworkNormalizeTests.bas, main.
Interface Normalize(source As ATNetworkReadResult,dateTextFormat As String,unlimitedKeyword As String) As ATNetworkReadResult.
- [ ] RED reader-to-normalizer-to-DSL integration plus standalone invalid/config/error/limits cases from spec.
- [ ] Implement raw copy/date classification/typed rows; full GREEN, commit.
- [ ] Fresh independent review; fix Important with RED/GREEN. Static7/7 gate9/9, clean-checkout full tests.
- [ ] Russian status/readiness/docs evidence; push dedicated branch, draft PR base feat/network-request-reader, attach and confirm latest-head CI. No merge/release.

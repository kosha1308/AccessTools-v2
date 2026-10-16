# Address declaration match Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL superpowers:executing-plans; native inline and one final fresh reviewer.

**Goal:** Exact upper/lower declaration comparison in existing local report.
**Architecture:** RAW snapshot, existing bounded tokenizer/token validator, distinct declaration keys and refs; no core changes.
**Tech Stack:** Offline VBA and supervised real Excel.
**Spec:** docs/superpowers/specs/2026-10-16-network-address-match-design.md.

## Global Constraints

Basefea4c22, baseline10430/0. Existing completed checks/writer algorithms unchanged.
No source/settings mutation. Limits10000rows/keys,65536text,200000steps,990matchissues,1000aggregateissues.
Groups/masks/cell errors SKIPPED comparison, invalid numeric tokens retained whole. No full-release claim.

## Review Focus

Equivalent coverage does not imply same IP/CIDR/range declaration.
Invalid tokens and stale prefix cannot disappear into shorter valid IPs.
One group/error anywhere prevents false missing/extra issues.
Merged/spacer refs and repeated declarations: preserve source, dedupe exact refs.
New issues plus existing reports can exceed cap: block without partial positives.

### Task1 — exact comparator and local integration

Files src/checks/ATNetworkAddressMatch.cls, ATNetworkLocalChecks.cls,
src/ui/modATReportLabels.bas, tests/vba/modATAddressMatchTests.bas, main;
integration expected outcome count9→10 only in existing tests/helper.
- [x] RED pair corpus and real-reader cases, invalid/missing/extra/groups/errors/limits, integration10outcomes.
- [x] Implement; real Excel full GREEN, commit. Fresh independent review; fix Important RED/GREEN.
- [x] Static7/7 gate9/9, original ReadOnly/hash, clean-checkout compile/full suite.
- [ ] Russian status/readiness, push draft PR base feat/network-table-report, attach and verify latest-head CI. No merge/release.

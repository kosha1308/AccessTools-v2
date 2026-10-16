# Network Domains Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: superpowers:executing-plans; native execution in this chat. User delegated technical decisions; no technical approval roundtrip.

**Goal:** Correct IP/port comparisons for the first NETWORK_ACCESS user version.
**Architecture:** Frozen interval payload, strict domain parsing, existing typed evaluator integration.
**Tech Stack:** Offline VBA and existing Excel COM runner.
**Spec:** docs/superpowers/specs/2026-10-16-network-domains-design.md.

## Global Constraints

IPv4 endpoints0..4294967295, ports0..65535, exact integral Double. No expansion
into individual IPs. Existing ATEvalBudget limits apply. Never copy private legacy
source or working inputs into public repo. Compiler catalog unchanged. Branch
feat/runtime-domains, based on verified runtime; no merge into main or prior PRs.

## Review Focus

Signed boundary2^31 (task1); trailing malformed token (task2); internal wildcard
explosion (task2); mutable domain payload (task3); evaluated fractional/invalid
port conversion hidden by valid list elements (task3). Each has explicit tests.

### Task 1 — interval algebra

Files: src/domains/ATIntervalRange.cls, ATIntervalSet.cls; tests/vba/modATIntervalTests.bas, test main.
Interfaces: range.Initialize(low As Double, high As Double), read-only Low/High;
set.Initialize(maximum As Double); Add(low As Double, high As Double, budget As ATEvalBudget);
Freeze(); read-only Frozen/Count/Item(index As Long) As ATIntervalRange;
Overlaps(other As ATIntervalSet,budget) As Boolean; IsSubsetOf(other,budget) As Boolean.
Reject mutation after Freeze and reinitialization; validate endpoint bounds.

- [ ] Write endpoints, adjacent/gap normalization, full-vs-partial, freeze and signed-boundary tests. Compare all subsets of 0..5 against Boolean-array oracle for overlap/coverage.
- [ ] Run Excel runner; observe missing classes. Implement sorted merging and two-pointer comparisons, charging steps and checking intermediate count before allocation.
- [ ] Run full suite, record actual evidence and commit.

### Task 2 — strict domain parsers

Files: src/domains/modATDomainParse.bas, modATDomainIP.bas; tests/vba/modATDomainParseTests.bas.
Interfaces: AT_ParseDomain(typeName As String, raw As String, node As ATAstNode, Optional budget As ATEvalBudget=Nothing) As ATEvalResult;
AT_ParseIPv4(text As String) As Double; AT_IPText(number As Double) As String;
AT_IPDeclarationsEqual(a As ATRuntimeValue,b As ATRuntimeValue,budget As ATEvalBudget) As Boolean.
Task3 value interface used by parser: InitializeDomain(typeName As String, intervals As ATIntervalSet, declarations As Collection).

- [ ] Write parser fixtures: /0,/32, normalized-host CIDR, compact/range/list, x masks including internal x; malformed tails/long octets/negative/overflow, port separators, empty; declaration strictness and invalid cannot compare; step/list/text limits.
- [ ] Run RED; implement full-token scanning with no regex, preserving declarations. Task2 and3 shared model may be introduced together to compile, ledger this interface dependency.
- [ ] Full GREEN and commit.

### Task 3 — immutable value and evaluator integration

Files: ATRuntimeValue.cls, modATRuntimeNormalize.bas, modATRuntimeDispatch.bas,
ATEvaluator.cls, modATRuntimeSets.bas; new src/domains/modATDomainEval.bas;
tests/vba/modATDomainEvalTests.bas; update old unsupported IP fixture to one still unsupported.
Interfaces: RuntimeValue.Domain As ATIntervalSet, DeclarationKeys As Collection;
AT_ConvertPortValue(value As ATRuntimeValue,node As ATAstNode,budget As ATEvalBudget) As ATEvalResult;
AT_ApplyDomain(node As ATAstNode,args As Collection,budget As ATEvalBudget) As ATEvalResult.
Domain payload shared only frozen; defensive declaration copy; no Scalar access.

- [ ] Write compiled field-field and contextual port expressions, numeric-call/list conversions; invalid fractional port and trailing invalid values ERROR; domain freeze test, lazy branch skip and full coverage vs overlap tests.
- [ ] Observe RED. Support domain type/signatures; apply recorded conversion after Visit evaluation, preserving non-OK propagation; dispatch to interval operators; keep row/source/regex unsupported.
- [ ] Full GREEN including1264 baseline, independent review/fixes with RED fixtures, tooling7/7, gate9/9, clean clone; commit, draft PR against runtime-core, attach and verify CI.

### Task 4 — usage milestone

- [ ] Update Russian status: domain kernel verified or exact remaining blocker. Do not claim first user version until NETWORK_ACCESS adapter, cell results, packaged XLAM and smoke comparison are verified. Record next concrete milestone.

Self-review: Task2 consumes Task3 domain payload; implement the immutable payload
as part of task2 after model/parser tests are written, then task3 wires runtime.
All remaining APIs are explicit; bounds/loop budgets and oracle tests own risky cases.

# Network policy catalog Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL superpowers:executing-plans. Native inline implementation and one fresh final review.

**Goal:** Apply validated protocols, message severity and routing to existing local checks.
**Architecture:** Immutable compiled catalog/decisions + wrapper runner; original checker and writer algorithms unchanged.
**Tech Stack:** Offline VBA, real Excel.
**Spec:** docs/superpowers/specs/2026-10-16-network-policy-catalog-design.md.

## Global Constraints

Base1369b65/suite10551. Four RAW tables,10000rows each; shared200000steps,
65536text,32767stored/expandedmessage. Existing1000issue/report cap. No source writes.
Explicit active/severity/routing flags; no invented protocol list or wildcard keys.

## Review Focus

Malformed active row must not silently disappear; inactive rows cannot add permission.
Severity override must not alter key matching or conceal blocked execution.
Inserted message text must not recursively expand placeholders.
Routing/user-error counters must remain distinct from full report error count.
Missing/duplicate policy, direct typed oversized input and repeated runs fail predictably.

### Task1 — validated immutable catalog and decisions

Create src/policy/ATNetworkPolicyCompiler.cls, ATNetworkPolicy.cls,
ATNetworkPolicyResult.cls, ATNetworkPolicyDecision.cls, modATPolicyText.bas;
tests/vba/modATPolicyCatalogTests.bas and main invocation.
Compile(protocols,problems,routing,commonText As ATEvalResult) As ATNetworkPolicyResult.
Result.State/Policy/Diagnostics guarded. Policy.Protocols copied Collection;
Decide(issue,key) As ATNetworkPolicyDecision; CommonText(key,fallback) As String.
- [x] RED tests valid/empty/inactive/duplicate/invalid/typed-budget, immutable,
 explicit flags/severity, original-key routing, nonrecursive templates and limits.
- [x] Implement minimal code, full Excel GREEN and commit.

### Task2 — apply policy around completed checks

Create ATNetworkPolicyChecks.cls, ATNetworkPolicyReport.cls, modATPolicyBindings.bas;
tests/vba/modATPolicyApplyTests.bas. Run(source,loadedPolicy) AsATNetworkPolicyReport.
Report exposes AllReport/Decisions/UserIssues/RequestIssues/SupersetIssues/UserErrorCount.
- [x] RED protocol membership, missing catalog skip, configured message/severity,
 original source refs, hidden routing, recalculated outcomes, blocked bypass,
 invalid policy no partial result and collection immutability.
- [x] Implement wrapper, full Excel GREEN, commit. Do not change existing checkers.
- [ ] Fresh review/fix, actual provided Rulebook read-only/hash, static7/7 gate9/9,
 clean-checkout default60s full suite; status/readiness, draft PR basePR15, attach/CI.

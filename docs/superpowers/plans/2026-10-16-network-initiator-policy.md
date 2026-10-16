# Initiator policy Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL superpowers:executing-plans. Native inline; one fresh final reviewer.

**Goal:** Independent admission, approvals and full-coverage IP policy with explicit no-policy action.
**Architecture:** Compiled selectors/regions +actor checker; compose13 raw outcomes, present once through existing catalog.
**Tech Stack:** Offline VBA, real Excel tests.
**Spec:** docs/superpowers/specs/2026-10-16-network-initiator-policy-design.md.

## Global Constraints

Base8ad6a2c,10660/0. Explicit network profile NoPolicyAction=ALLOW per user.
10000rows/table,200000sharedsteps including text,65536text,10000intervals.
No source writes/core changes. Published blank-defaults apply only to IP table.

## Review Focus

Exact email replaces domain; negative selectors never accidentally admit.
Union coverage +deny precedence, invalid policies do not partially parse.
Scope boundaries and lower unknown-direction addresses cannot bypass restrictions.
Approval is not a user-input error; blocked execution cannot disappear in routing.
Empty policies, explicit NoPolicyAction, native account refC10, immutable snapshots.

### Task1 — compile and select actor policy

Create src/policy/ATInitiatorSelector.cls, ATInitiatorPolicy.cls,
ATInitiatorPolicyResult.cls, ATInitiatorPolicyCompiler.cls, ATInitiatorSelection.cls,
modATInitiatorText.bas, modATNetworkPolicyProfile.bas;
tests/vba/modATInitiatorPolicyTests.bas, main invocation.
Compile(initiators,ipRules AsATEvalResult,noPolicyAction AsString) AsATInitiatorPolicyResult.
Policy.ForEmail(email,budget) AsATInitiatorSelection with IdentityReport(ref),
Evaluate(expression AsATRuntimeValue,scope,budget) AsATEvalResult, AccountForbidden,
HasDirectionalRules; selections immutable. Source primitives use existing interval sets.
- [ ] RED selector/identity/approvals and coverage corpus: exact/domain priority,
 exclusions, wildcard boundaries, deny/NOT, multi-row union, scope, no-policy,
 malformed/inactive/error/budget and snapshots; observe failure.
- [ ] Implement, real Excel full GREEN, commit.

### Task2 — checks and one-pass composition

Create ATNetworkInitiatorChecks.cls, ATNetworkAccessChecks.cls;
modify ATNetworkPolicyChecks (extract Apply), policy bindings/default approval and labels;
tests/vba/modATInitiatorApplyTests.bas.
Actor.Run(source,loadedActor) AsATCheckReport3 outcomes.
Access.Run(source,loadedCatalog,loadedActor) AsATNetworkPolicyReport13 outcomes.
- [ ] RED source/destination denied refs, lower actuals/unknown direction,
 accountC10, mixed invalid inputs, missing email, admission/approval independence,
 unchanged local10, routing/one-pass messages and failures.
- [ ] Implement, full GREEN; fresh independent review/fix.
- [ ] Read-only original/hash, static7/7 gate9/9 clean-copy default60s full suite;
 status/readiness, draft PR basePR16, attach/verify CI. No merge/release.

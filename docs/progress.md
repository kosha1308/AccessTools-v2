# Stage 1 progress

## Runtime continuation 2026-10-16

User delegated technical decisions and asked to continue after the Excel restriction
was disclosed. Ruling: proceed with code and written regression tests while preserving
explicit unverified status; do not claim RED/GREEN or stage acceptance. No sandbox
escalation, Trust Center change or alternate execution bypass was attempted.
Branch feat/runtime-core isolates implementation from the verified compiler branch.
Runtime model, explicit normalization, typed traversal, lazy logic/state inspectors,
named bindings, scalar/list handlers, budgets and trace are drafted. Binding IsDefault
is explicit metadata, separate from diagnostic SourceId. Snapshot safely shares only
internally immutable results; field validation checks budgets before returned copies.
Independent static review reported four issues (trace primary failure loss, nested
supplied-argument scope, DATE time-fraction bypass, copying before list limits).
Corrections and dedicated regression fixtures are written; execution remains pending.
Observed checks: 56-component structural validation PASS; static tooling 7 PASS / 0 FAIL;
git diff whitespace check PASS. Excel gate fails at COM activation E_ACCESSDENIED
80070005 before compilation. No new VBA test totals or compilation success asserted.
Detailed status and remaining acceptance: docs/runtime-core-status.md.

## Historical stage 1 evidence

Plan: docs/superpowers/plans/2026-10-15-stage1-plan.md
Design and plan approved in chat on 2026-10-15.

Execution: inline. A dedicated fresh clone and feature branch isolate changes; no extra worktree is necessary (main remains untouched).
Baseline: only README and approved design/plan; no existing tests.
Environment: Excel COM and VBProject access verified in a separate Excel instance. Sandbox COM activation requires escalation; no Trust Center settings changed.
Pre-flight: task 1 runner feeds all VBA suites; task 2 registry/diagnostics feed lexer and compiler; lexer tokens feed parser; AST feeds serializer/resolver/type checker. Interfaces will be versioned in docs/dsl/contracts.md. No interface conflict found.
Task 1: RED -> GREEN: missing validator rejected; 7 malformed/valid tooling fixtures pass. Excel RED failed compilation with missing version constants; after implementation Compile VBAProject PASS, VBA tests 2 PASS / 0 FAIL.
Task 1: Excel runner fixtures: passing result, intentional assertion failure, syntax error in unexecuted procedure, and zero-test result are independently exercised. Process-scoped ExecutionPolicy Bypass is used only for developer scripts (no settings persisted).
Task 1: Ruling: use dedicated fresh clone/feature branch instead of an additional worktree. Cost if wrong: workspace relocation only; main and old AccessTools are unaffected.
Task 1: complete (fd5d368): static tooling 7/7, Excel gate integration 4/4, explicit Compile VBAProject and VBA tests verified. Missing Excel/denied Trust Center were not simulated by changing system settings; worker exceptions fail closed, actual sandbox denial was observed during prerequisite probing.
Task 2: RED -> GREEN: missing classes failed compile; registry then compiled and passed 14 assertions. Mutation of registered signatures and unknown return type each reproduced a failure; both fixed, full suite 16 PASS / 0 FAIL with Compile VBAProject PASS.
Task 2: complete: foundational registry/diagnostic contracts. Bootstrap is deliberately a subset; full appendix-N catalog is added with parser/compiler consumers, not advertised as implemented.
Task 3: complete: RED missing lexer failed compile; GREEN Compile VBAProject PASS, full suite 29 PASS / 0 FAIL. Unicode/escaping, CRLF source spans, decimals, references and malformed suffix covered.
Task 4: complete: RED missing AST/parser failed compile; GREEN parser/serializer suite covered precedence, contextual/full equivalence, nested aliases, calls/lists, malformed input and round-trips. Additional string-as-keyword case reproduced a real failure and was fixed by token-kind-aware matching. Final full suite 57 PASS / 0 FAIL, Compile VBAProject PASS.
Task 5: compiler foundation implemented: context/resolver/type checker, typed named parameters/defaults, cycles, metadata sources/dependencies, deprecated warnings and UsedBy edges. RED missing compiler failed compile; initial GREEN 84 assertions. Broader full-catalog acceptance is open, see docs/stage1-status.md.
Task 6: limits and seeded malformed-input/property coverage implemented. Boundaries and exceeded limits tested; appendix compiler samples covered. Full phase acceptance is not asserted.
Final review: independent reviewer reported 5 important issues; fixed default-recursion guard, schema-type validation, integer JSON counters, registry syntax resolution including multiword operators, and duplicate syntax registration.
Final fixes: multiword operator test RED -> GREEN; string counter integration RED reproduced garbage PASS -> GREEN (6/6 integration cases). Default-cycle/schema-type regressions and syntax ambiguity regression pass. Deprecated-warning RED exposed an unsafe test assertion/worker failure; safe assertion plus warning implementation is GREEN.
Verification: explicit Compile VBAProject PASS; VBA 787 PASS / 0 FAIL; tooling fixtures 7/7; Excel integration 6/6. Malformed-input generator has a fixed seed and does not execute DSL as code.
Ruling: grammar/compiler foundation milestone is published with stage acceptance OPEN; unsupported source lookup/legacy import are diagnosed, not silently accepted. Cost if mistaken: premature adoption would require withheld runtime functions; no production pipeline is connected.
Ruling: executable appendix cases live in tests/vba/modATAppendixTests.bas instead of unused JSON fixtures. Cost if mistaken: future fixture tooling may need format conversion, not semantic changes.
Clean-checkout verification: 52b0976 cloned locally without hardlinks; explicit Compile VBAProject PASS, VBA 787 PASS / 0 FAIL, static fixture suite 7/7. No old AccessTools checkout or generated artifacts were required.
Independent-review smaller gaps: deprecated warnings and duplicate parameter checks implemented; broad catalog/golden/domain-normalizer coverage remains explicitly listed as stage acceptance work, not deferred cosmetic findings.
## Compiler acceptance continuation (2026-10-15)
Catalog continuation: RED 39 missing/incorrect cases -> GREEN 831 assertions after common catalog/parser/literal changes. Fixed a test-only syntax-name collision introduced when MATCHES/IS NULL became built-ins. Added test-stage tracing and an outer VBA assertion boundary so unexpected test exceptions fail as JSON instead of waiting on a VBE modal dialog.
Independent catalog review: singleton contextual collections, named port argument/default validation and chained alias negation reproduced (8 FAIL), then fixed; full Compile VBAProject PASS, 839 assertions PASS / 0 FAIL.
Catalog declarations include text, set/list, state-inspection, BETWEEN, NOT IN, numeric/aggregate/default/time functions. Their evaluator bindings remain metadata; no Runtime evaluator was enabled. NUMBER canonicalization is string-based and preserves exact decimal spelling without floating-point rounding.
Typed source schemas and column/key checks, hygienic DOCUMENT/ROW named definitions and defaults, selected signatures, parsed argument bindings, schema/DSL versions and exact diagnostic provenance completed. Regression milestones: 853, 874, 879, 882, 886 assertions; every registry operator/function example now compiles (947), followed by SET/list and appendix contracts (953).
Catalog snapshot is generated from the runtime registry: 76 entries, including 61 operator/function/quantifier examples. Missing or drifting snapshots fail the gate; regeneration requires explicit -UpdateCatalog.
Final development verification: explicit Compile VBAProject PASS; VBA 953 PASS / 0 FAIL; static tooling 7/7; Excel integration 9/9, including missing/drifting/matching catalog fixtures. Normal Test.ps1 run verified the committed catalog without updating it.
Final independent review: no important blocker. Invalid NOT_EMPTY and placeholder examples fixed and verified across all examples. Deferred to evaluator/explain: nested deprecation-warning SourceId defaults to RULE; directly typed lookup/source literals, BIND and synthetic contextual LIST nodes can omit AST SourceId. Error diagnostic provenance is covered and retained.
Rulings: port 0 is a valid unsigned numeric literal (business restrictions belong to policy); source LOOKUP uses declared metadata only and opens no files; named defaults resolve in definition scope with prior parameters. Stage closure is compiler profile 1.0.0 only; no evaluator, real-request cutover or old AccessTools changes.
Clean-checkout acceptance: commit 3507720 cloned locally without hardlinks into AccessTools-v2-stage1-verify. Explicit Compile VBAProject PASS, VBA 953 PASS / 0 FAIL, catalog snapshot matches, static fixtures 7/7. No previous generated files or old AccessTools checkout required.
Runtime verification restored: escalation is now available; Excel launched successfully. Compile failures localized by new runner output: cross-instance Private RetainField and VBA reserved local/scale names. Fixed Friend access with frozen guard, localFrame/factor names. First full execution: 1236 PASS / 12 FAIL. Exact LIST overloads lost to PORT_SET conversions, and generic table T bound to empty type; two-pass overload search plus explicit generic binding fixed these: 1258 PASS / 0 FAIL.
Final independent review found invalid trailing list elements could yield a Boolean and AVG could underflow nonzero to zero. Added reproducing fixtures: 1260 PASS / 3 FAIL. Eager budgeted list usability validation and AVG range guard yield Compile PASS, 1264 PASS / 0 FAIL (artifact excel-250143f025454ad6ae4c93cf049520f6). Earlier four static-review corrections are covered by executed fixtures. Static tooling 7/7, Excel gate integration 9/9 including compile-error diagnostics. No important findings deferred.

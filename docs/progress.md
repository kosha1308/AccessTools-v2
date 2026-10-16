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
Clean-checkout acceptance: 0412cee cloned without hardlinks into AccessTools-v2-runtime-verify. Compile VBAProject PASS, 1264 PASS / 0 FAIL; catalog snapshot matches; static fixtures7/7. Report excel-15c20ebe9caa4fcda4935a5c73901f0c. Remote API publication history reconciled only after exact baseline tree equality was checked; no force push and no merge into compiler/main.
## Network domains 2026-10-16
User requested continuing toward usable v2 for network connections. Clarified first-use requirement: full familiar workflow with request checks, Superset and coordinator summary. Milestone recorded in docs/first-use-network-access.md; do not equate domain kernel with user-ready release.
Read Drive spec v0.13 and legacy source as semantic reference only; no private source/data copied into public v2. Task1 missing classes RED -> CompilePASS9461/0, with all4096 pairs of finite sets checking both overlap/full coverage. Task2 missing parser RED -> CompilePASS9507/0, including strict malformed tokens, compact/range/x/CIDR and port separators/bounds; domain payload shared only frozen.
Task3 unsupported operations RED20FAIL -> runtime wiring. Three remaining expectation mistakes corrected: COVER A subset R versus CONTAINS R subset A. Result9553/0. Independent review found compact left-range continuation missing; regression1FAIL reproduced, prefix fix plus valid ABS/SUM/list-field/computed-default conversion tests ->9563/0 (excel-e7c924a99c01404d9167fc95e097093f). No important/minor review findings deferred. Static fixtures7/7, Excel gate9/9. Unchanged catalog snapshot matches.
Ruling: kernel port0 stays unsigned-valid; NETWORK_ACCESS adapter may impose business range1..65535. IP reverse endpoints normalize as legacy interval layer; port reverse ranges reject. Coverage equality and strict CIDR declaration equality remain separate helpers. Domain invalid data never becomes successful comparison. Masks with nontrailing x may hit explicit work/interval limits instead of expanding unboundedly.
Clean-checkout domains acceptance: fc972df cloned without hardlinks into AccessTools-v2-domains-verify. CompilePASS9563/0, catalogmatch, static fixtures7/7; reportexcel-3ad9f0ddaa694637bf1f1ff0601fdf1c. No previous generated artifacts or legacy checkout required.
## Row runtime 2026-10-16
Full-first-use scope remains NETWORK_ACCESS+Superset+coordinator, no technical approval roundtrip (delegated). Branch feat/runtime-rows based729eb05. Compiled FieldId/TableId/RowAlias/NamedScope metadata: missing API RED→9571/0. Frozen row model and collections: missing class RED→9580/0. Scoped traversal and row/cell trace: missing trace API RED; after implementation9660/0 including row predicates, empty laws, nonOK states, correlated IP, filters and budgets.
Six old unsupported-row negative fixtures now target row operators over unsupported SOURCE, because row operators themselves are supported. This preserves negative source behavior while adding positive row coverage; no parser/catalog reduction.
Independent review found no concrete bug; three actionable coverage gaps closed with oversized unused fields, DOCUMENT default alias shadow, missing fields and evaluator reuse after nested error/skip. Full CompilePASS9679/0, reportexcel-c41f3ebcb1774df78b5ec28cc4ffb105; catalogmatch. Static7/7, gate9/9. No review findings deferred. Rows remain in-memory; Excel extraction/Superset/coordinator are next.
Clean-checkout rows acceptance: 3ea739f cloned locally without hardlinks into AccessTools-v2-rows-verify. CompilePASS9679/0, catalogmatch, static7/7; reportexcel-75c1127ff3e74a6286bdabe92a0f6fc2. No old artifacts or private legacy checkout needed.

## Excel table extraction 2026-10-16

Branch feat/network-extraction based bf00172. Generic external header/usage descriptions,
sheet-bound copied layout and bounded Value2 reader. Raw TEXT, immutable physical rows,
merged-anchor cell references, CVErr fields, visual emptiness, DIRECT/INHERITED/NONE
participation. Last-value scan ignores formatting/formula-empty; physical/cell/find limits.
Locator missing-class RED ->9694/0; reader missing-class RED ->9721/0.
Fresh read-only reviewer found two Important merge bugs: final vertical extent truncated,
horizontal merge duplicated independent fields. Real-Excel repro9721/4 ->9728/0 after fixes.
Additional boundary/owned-horizontal/closed-workbook coverage ->CompilePASS9733/0;
reportexcel-ab3b8114e755480e8c9e6d33c9010241. Catalogmatch, static7/7, gate9/9.
Gate initially lacked COM permission in ordinary sandbox; rerun with existing Excel
authorization passed9/9. No Trust Center/settings changes.
Review observation ruled: layout validity limited to one synchronous unchanged-structure
read, documented. Full fingerprint belongs to published profile stage. Other declined
items (full profile, single fields, rulebook, Superset/coordinator/UI) are explicit pending
first-use requirements, not silently dropped. No review correctness findings deferred.
Clean-checkout acceptance: e23614e cloned without hardlinks into AccessTools-v2-extraction-verify.
CompilePASS9733/0, catalogmatch and static7/7; reportexcel-f960b4c08d56499285469391d3fd9660.
Build uses only this repository, without old artifacts/private sources. Local clone required
existing elevated process permission after ordinary sandbox blocked Git's helper process.

## Document fields 2026-10-16

Branch feat/document-fields based1401791. Fresh baseline9733/0, reportexcel-a2b88bcc7ede46a58d3a9ba675bc3fae.
Immutable external field specs/results and bounded locator with signed offsets from merged
boundaries; raw TEXT and Found/MISSING/EMPTY/ERROR metadata. Shared source-cell helper.
Missing-class RED ->9774/0; optional document AddField CellRef missing-API RED ->9787/0.
Document Snapshot and evaluation trace preserve source cell, distinct from active row fields;
ERROR, lazy OR and metadata text budgets covered.
Fresh independent reviewer: one Important partial merged anchor crossing left/top search edge.
Real Excel9788/2 RED ->9792/0 with bounded leading-edge metadata validation. Added metadata
limit/reuse/frozen spec coverage ->CompilePASS9796/0, reportexcel-d2b1d72852b04ba791a94f6219a681a6.
Static7/7, gate9/9, catalogmatch. Boundary merge probing<=10000, unchanged plain edges skipped
in one query. Ruling: all declined production profile/fingerprint/normalizers/rulebook/Superset/
coordinator/UI items remain explicit pending first-use requirements; no correctness findings deferred.
Clean-checkout fields acceptance:1a97f21 cloned without hardlinks into AccessTools-v2-fields-verify;
CompilePASS9796/0, catalogmatch, static7/7, reportexcel-8c40967230df469bbdc205a7f2c64558.
No legacy source or old generated artifacts required.

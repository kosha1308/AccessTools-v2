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

## Revised TZ / Superset enrichment 2026-10-16

Fetched current and immediately previous Drive revisions of v0.13. Previous matches cached
162214 chars; current163176 chars modified2026-10-16T10:34:07.429Z. Delta: SKUF group and
support team from actual visible linked Superset rows, header whitespace tolerance, expanded
real-app smoke acceptance. Old baseline85/0/0 is explicitly not v2 evidence. Current local TZ
cache updated, previous preserved outside public repo. No private old code/data copied.
Ruling: implement revised enrichment contract now; pending template fingerprint remains scheduled.
Branch feat/superset-enrichment based0aef522. Fresh baseline9796/0 reportexcel-1f8b2723714d4914ab3fa34b1719bab2.
Text-set/bucket missing-class RED ->9809/0, reportexcel-8dc3cc540ab84b5e8823168c3e6c47cb.
Visible linked adapter missing-class RED ->9834/0; expanded ids/schema/Cyrillic/long-prefix
coverage ->CompilePASS9845/0, reportexcel-da4bf09d1a7249998cc068d00a735f61. Source snapshot
copied once; number canonicalization linear. Static7/7, gate9/9, catalogmatch.
Fresh independent review: no requested changes; trim semantics clarified in spec. Declined
items are explicit pending pipeline/release scope; combined maximum-limit performance not
measured, bounded work limit tested. No correctness findings deferred.
Clean-checkout enrichment acceptance:69e03d2 cloned without hardlinks into AccessTools-v2-enrichment-verify.
CompilePASS9845/0, catalogmatch, static7/7; reportexcel-e9a8f1184d4a4fc19e51b8fb91314262.
No old artifacts or private runtime checkout required.

## Template structural gate / confirmed reference 2026-10-16

Branch feat/template-gate basedecdee21. Fresh baseline9845/0 reportexcel-55760ebc60da4cd596b472ea16bcea48.
Profile/decision missing-class RED ->9869/0; guarded reader missing-class RED ->9874/0.
User supplied reference location during execution: Drive work/Шаблоны заявок, XLSX
1An-gcbJkeZl0u4KsvmIAnUgSA4b5jiF9 (16181bytes). Original kept outside repo. Inspected
OOXML structure with bundled Python, then checked real XLSX ReadOnly in owned supervised Excel.
Ruling: concrete profile mirrors unprotected reference; no invented locking requirement.
Grouped anchors added to avoid treating user text as an isolated dynamic lower header.
Concrete factory missing RED ->9880/0. Actual reference MATCH NETWORK_ACCESS@3.0-profile1;
SHA256 unchanged21684A4A985FD2057193B557E9D09C817C0899CE09DF7C697EF45025BC6714A7.
Fresh independent review: stale absolute layout could expose moved header as data. Reproduced
9881/2 RED ->9883/0 with guard accepting spec/re-locating after Verify. No external callers
of new draft API; spec/plan/tests updated. Extended config/count/copied group/guard cases
->CompilePASS9889/0 plus actual XLSX MATCH, reportexcel-d04e7e2da1f34a0cba720e38d134f288.
Static7/7, gate9/9, catalogmatch. No review correctness findings deferred. Full app dispatcher,
normalization/checks/Admin publisher/IP matcher/cache/coordinator writer/XLAM remain pending.
Clean-checkout template gate acceptance:4693424 cloned without hardlinks into AccessTools-v2-template-verify.
CompilePASS9889/0, catalogmatch, static7/7; reportexcel-afcbac1566b34cc487d141a78fdab562.
Normal test suite reproduces without original private XLSX; separate real-source read-only
acceptance and unchanged hash are recorded above.

## Unified request reader 2026-10-16

Branch feat/network-request-reader based f3ff679. Task1 native scalar preservation/shared
metadata RED ->9898/0. Task2 missing classes RED, then9938/2: DSL test compared present
text with EMPTY literal (RUNTIME_VALUE_UNUSABLE). Test corrected to LEN >0, no runtime
semantics changed ->CompilePASS9941/0, reportexcel-b2582e1066004f22951711e01e06bff6.
Fresh independent review of f3ff679..021a145 found no actionable issues.
Actual source XLSX opened ReadOnly, unified extraction OK|TECH=0|ADDR=0|DB='Лист1'!C10,
full9941/0 reportexcel-a914b7b0fcb94b6fb4e9d5efd3e5495b. Source SHA256 unchanged
21684A4A985FD2057193B557E9D09C817C0899CE09DF7C697EF45025BC6714A7.
Static tooling7/7. Raw snapshot only; normalization/business checks/full release remain pending.
Clean checkout e2bbfd1 cloned without hardlinks to AccessTools-v2-request-verify: CompilePASS9941/0, static7/7, reportexcel-a484ff4fc4364d4790d4966d64a23980. Excel gate integration9/9. No private reference required for normal tests.

## Network normalization 2026-10-16

Base7ccdd27. Date adapter missing-class RED ->CompilePASS10169/0
reportexcel-7e960182ceec4ef794d1aabd2ec5ebd7. Typed adapter missing-class RED,
reserved VBA identifier/test fixture merge and unregistered test function fixed;
no core behavior changed ->CompilePASS10250/0 reportexcel-b9fd26b171464bd0867a5af224555902.
Fresh independent review found inconsistent keyword canonicalization and acceptance
of raw system failures. Added reproductions10254/8 ->canonical validation + raw
failure allowlist(EXCEL_CELL_ERROR only) ->CompilePASS10265/0
reportexcel-df9d5f08c31f40e2bb6117ba1293b271. Both findings resolved, none deferred.
Actual source ReadOnly full pipeline OK|TECH=0|ADDR=0|UNLIMITED=True; unchangedSHA256
21684A4A985FD2057193B557E9D09C817C0899CE09DF7C697EF45025BC6714A7,
reportexcel-4fa40765b87a42ff9fc143bbc1224862. Static7/7, gate9/9.
Ruling: clarified shared runtime budget200000steps/65536text,10000rows;1mcell limit
belongs to reader. Date formats/keyword explicit; no EMAIL/enum/protocol allowlist
or revoke/required-fields policy invented. Full release remains pending.
Clean checkout133f603 cloned without hardlinks to AccessTools-v2-normalization-verify: CompilePASS10265/0, static7/7, reportexcel-85b8fe0081d94d5d9e5635473728e9c8. No private source file required for standard suite.

## Six document checks / issue report 2026-10-16

Basea462767. Read-only legacy semantic reference main d9aa9b88f9193c0d2935d99c5cc8ecca3865687c
confirms4required initiator fields, optional contact, request aliases, IEP yes,
justification warning, revoke-ignore and embedded DMY dates. No private code copied.
Task1 DMY_TEXT RED10266/21 ->CompilePASS10316/0 reportexcel-11e9a7b312e94564a4f027a7b54273d3.
Task2 missing-class RED ->CompilePASS10381/0 reportexcel-118e6a29e51a464e9eedc71e840e3b86.
Fresh independent reviewa462767..858bd71: no actionable findings. Report immutable,
scopeNETWORK_DOCUMENT, six local checks only; no full-app PASS claim.
Actual original XLSX opened ReadOnly: STATE=0|CHECKS=6|ERRORS=4|WARNINGS=1,
full10381/0 reportexcel-c1e999ffef3b42cb904c8eece4f1bc47. Source SHA256 unchanged
21684A4A985FD2057193B557E9D09C817C0899CE09DF7C697EF45025BC6714A7. Static7/7.
Full table checks, source lookups, dispatcher/Rulebook UI, Superset and coordinator writer/package remain pending.
Clean checkoutbec2fd4 cloned without hardlinks into AccessTools-v2-document-checks-verify: CompilePASS10381/0, static7/7, reportexcel-d8ed8e2ee472479eb1e5b3f26b17009c. Standard suite needs no private original. Excel gate integration9/9.

## Table input checks + Russian workbook report 2026-10-16

Basef5e3204. Task1 missing-class RED; reserved local identifier renamed. FirstGREEN10403/0.
Separator-only protocol reproduction10403/1 ->explicit no-token rejection ->10404/0
reportexcel-a353bd6288bb4c4cbb5c159bac2ab357. Task2 missing writer class RED ->10424/0
reportexcel-35e448f10bd2461a87ae496d97b03f8e. Native hyperlink Follow verified actual
owned unsaved source/A5. Writer no original mutation, literal formula-like text,
foreign refs, copied report, events restoration/repeat verified.
Fresh whole-branch review found inappropriate IP tokenizer joining hyphenated protocol
words. Visual QA found lost opening apostrophe in ref text. Reproduced10426/4 ->
protocol-only splitter + literal prefix escaping ->CompilePASS10430/0. No core parser changes.
Actual original ReadOnly: STATE=0|CHECKS=9|ERRORS=6|WARNINGS=2|LINKS=8; unchangedSHA256
21684A4A985FD2057193B557E9D09C817C0899CE09DF7C697EF45025BC6714A7.
Final acceptance/sample/report-preview artifacts:excel-91d53946b05b46d49a3a99d08b47c75c.
Two-page native Excel print export rendered with bundled pypdfium2 and inspected;
readable, no clipping, full Russian labels, scope disclaimer. Private QA exports only.
Static tooling7/7. Review finding resolved; no correctness finding deferred.
Rulings: native VBA writer is application implementation tested in required Excel engine;
print/XLSX examples are QA intermediates, not standalone artifact-tool deliverables.
Conservative group/manual and explicit protocol catalog semantics retained. Exact
TOP-LOWER, Rulebook/source loaders, Superset/coordinator and app package still pending.
Clean checkoutf91d83d cloned without hardlinks to AccessTools-v2-table-report-verify: CompilePASS10430/0 under standard60s gate, static7/7; reportexcel-2f905fcb182f4b4c955fc95265a71a99. Excel gate integration9/9. Standard suite needs no private original.

## Exact address declaration match 2026-10-16

Basefea4c22; added NET.ADDRESS_MATCH as tenth local outcome. Declaration keys
preserve IP/CIDR/range distinctions and invalid tokens; groups/masks/errors skip
comparison. Missing/extra issues retain exact source refs; limits fail closed.
First GREEN10483/0; numeric invalid tail with x RED10484/1 ->GREEN10485/0.
Fresh independent review found missing typed-input text budget; RED10486/1 ->
CheckText through ResolveField budget ->CompilePASS10488/0.
Original ReadOnly STATE=0|CHECKS=10|ERRORS=6|WARNINGS=2|LINKS=8;
unchangedSHA256 21684A4A985FD2057193B557E9D09C817C0899CE09DF7C697EF45025BC6714A7.
Reportexcel-9bb89ef98b3f47a7a50cfa641641ee1d. Static7/7, Excel gate9/9.
User can provide empty Superset/Rulebook schema files; no real business data
required. Source readers, matching/coordinator and full user package remain pending.
Clean checkout0daf296: CompilePASS10488/0 under standard60s; static7/7, reportexcel-3ee05fd0e294460794d08605fa60d578.

## Published source schemas / read-only reader 2026-10-16

Base86e2f29. Supplied empty Superset +7 Rulebook data sheets mapped by real headers.
Source reader snapshots TEXT/EMPTY/EXCEL_CELL_ERROR with row IDs/refs, including
hidden and filtered rows. Unnamed export index ignored; no active/policy inference.
Limits100000body rows/1m rectangular cells/100000discovery work. Source-only changes.
TDD missing class RED; hidden-row omission reproduced, filter regression10550/1.
Disposable Excel probe confirmed Find/End ignore filtered row while SpecialCells
constants/formulas include it. Bounded area scan ->CompilePASS10551/0.
Actual two books ReadOnly counts0/3/0/0/48/3/48/0 and SHA256 unchanged;
reportexcel-96064cbbf3414ee0b6e3c3c6e8a5a508. Static7/7, Excel gate9/9.
Fresh independent review86e2f29..ea5ae65: no actionable Critical/Important findings.
RAW source readers complete; policy/catalog/messages/routing application,
Superset matcher/hostname refinement, coordinator output and full package pending.
Clean checkoutea5ae65: CompilePASS10551/0 default60s, static7/7; reportexcel-e17a3d029c6f4ccba6f12807a971ae5e. Standard suite requires no private source books.

## Validated Rulebook catalog / local policy application 2026-10-16

Base1369b65. Four RAW tables compiled: PROTOCOLS/PROBLEMS/ROUTING/COMMON_TEXT.
Immutable protocol list, explicit flags, scoped original-severity keys, custom
user/coordinator messages and nonrecursive template expansion. Wrapper invokes
existing10checks unchanged, recalculates severity outcomes, preserves skipped/
blocked, exposes routed channels and separateUserErrorCount. No source annotations.
Task1 missing-class RED ->10620/0; Task2 missing-class RED ->10639/0.
Fresh review found protocol compiler/consumer charset/count mismatch; verified
existing contract and reproduced. Added1000/1001 and direct10001-row boundaries.
Diagnostics guard + config-error classification RED10654/2 ->fixed; shared text
budget RED10657/1 ->Step consumed text lengths ->CompilePASS10660/0.
Real provided Rulebook/request ReadOnly: POLICY0/PROTOCOLS0/CHECKS10/ERRORS6/
WARNINGS2/USER_ERRORS6/USER_ISSUES8. Both hashes unchanged; reportexcel-
439609ba75994b08b6af611e953ccf78. Static7/7 gate9/9. No review finding deferred.
General initiator/IP/hostname/port/service rule evaluation, source matcher,
coordinator UI/outputs and user package remain pending. AllReport is not a
filtered user report; future UI must consume channel data/UserErrorCount explicitly.
Clean checkout1c6d8de: CompilePASS10660/0 default60s, static7/7; reportexcel-4fe55b829f664bd79f0e50015e90c92c. Standard suite needs no private sources.

## Initiator admission, approvals and IP access 2026-10-16

Base8ad6a2c/suite10660. User explicitly confirmed NoPolicyAction=ALLOW when no
personal email/domain IP rules. Published NETWORK profile constant, compiler action
required. INITIATORS +INITIATOR_IP typed RAW compile to immutable selector/selection.
Exact-email IP policy replaces domain globally, scopes ANY/SOURCE/DESTINATION,
allow union full coverage and no deny overlap, embeddedНЕ:/NOT:, x/х masks.
Published header defaults apply to IP table only; empty allow remains emptyset.
Identity admission independent from approvals and IP; account forbidden atC10.
Task1 RED ->10723/0; Task2 RED/reserved local variable fixed, fixture merge clear
corrected ->10734/0. Composite13raw checks thenApply once; approval fallback request
only/non-user, explicit catalog can override. Existing local10/core/writer unchanged.
Fresh review found malformed selectors and unparsed IP PASS when no policy;
reproduced10734/7 ->strict config domains/leading@, manual on unparsed ->10746/0.
Whole-form email checker unchanged. Original Rulebook/request ReadOnly ACTOR0/
CHECKS13/ERRORS6/WARNINGS2/ACTOR_SKIPPED3; both SHA256 unchanged; reportexcel-
24f24e7edf0a4e9ab530bde7f115d5bf. Static7/7 gate9/9. No review finding deferred.
Skipped actor reasons now explicit Russian strings. General IP/hostname/port/service
rules, Superset index/matcher, coordinator UI/outputs/instructions and package pending.
Clean checkoutbd41767: CompilePASS10746/0 default60s static7/7; reportexcel-80ac01f7972e468e835ce26997bb7035. ExcelGate9/9; standard suite no private sources.

## Compiled general Rulebook engine 2026-10-16

Base9027ece/PR17. New immutable RULES compiler/policy and Technical/Superset
evaluators, no changes to existing13 checks. Scoped IP overlap trigger, full
port coverage per individual rule, exact protocol/service conditions, hostname
specificity exact>prefix>contains then length among base-eligible rules.
Equal best ties plus hostname-free rules retained; conflicting service/group
outputs explicitly flagged. All seven outputs and source refs preserved.
No partial result on error/limits/incomplete context. Instruction-only output valid.
BlankACTIVE inactive per actual header, blankseverityINFO per legacy.
Current legacy semantic reference d9aa9b8 only; no legacy implementation copied.
RED10746/1 ->10780/0, extra budget edge RED10801/2 ->10803/0.
Fresh review unknownprotocol/malformedhostname plus current strict pipe-separator
regressions RED10803/11 ->10814/0, explicit Excel compilePASS. Closed both Important
findings; no findings deferred. Private reference RulebookReadOnly3/0 unchangedhash
52372F64E26ED59E06684F8A3CAE17E41131DED95F0EC4254B22E46BD4EFE920.
Static7/7, ExcelGate9/9. Runner integration, Superset shared matcher/cache and
final coordinator workbook/UI/package remain pending; this is not user-ready v2.
Clean checkout b2d1ca4: CompilePASS10814/0 default60s, static7/7, reportexcel-
de731ec2b8af4906a633ecc338e528e5. Final actual RulebookReadOnly3/0 reportexcel-
516ef4a7ca904ceb8f8e275943470410, SHA256unchanged.

## Technical Rulebook check integration 2026-10-16

Base648ea41/PR18. ATNetworkRuleChecks adapter, NET.IP_EXCEPTIONS messages with
full technical context, outputs/instructions/source refs, explicit conflicts.
Literal request protocol splitter/dedup reuses existing delimiter semantics.
Unknown/Excel-error contexts incomplete; empty table notPASS, full ports only.
New RunWithRules preserves original Run13 and adds14th outcome, Russian Excel
labels/writer path. Rule hits/conflicts default coordinator-only/non-user errors;
explicit catalog/routes override. Message/component limits replace partial new
hits with BLOCKED, independent previous13 remain.
RED10175/2 ->10836/1: remaining failure was fixture Excel coercion443,9000 to4439000,
diagnosed and text format fixed; no product parser weakening ->10841/0.
Fresh review2 Important reproduced10841/2: severity downgrade could hide incomplete,
custom message expansion could erase13. Immutable Incomplete optionaldefaultFalse,
independent rule presentation and report/decision merge preserve completed results.
Reverse severity upgrade regression10844/1 ->10845/0 CompilePASS. Each final issue
formatted once; previous formatted issues never re-enter template expansion.
Original 13-check algorithms unchanged; defaultFalse preserves old outcome behavior.
Actual original empty template and Rulebook ReadOnly6/0: first13 statuses/errors
identical, new14, writer separate workbook; bothSHA256 unchanged.
Scope stays NETWORK_ACCESS_PARTIAL_POLICY. Hostname/Superset context, shared
source index/matcher/cache, final coordinator workbook and package still pending.
Final clean checkouta098fea: CompilePASS10845/0 default60s, static7/7,
reportexcel-8dee5743a93d46bf8db0066fce7a05e7. ExcelGate9/9.
Final actual ReadOnly6/0 reportexcel-8c3ca339d69b48d0a836d9152853f633,
bothSHA256unchanged. Standard suite no private inputs.

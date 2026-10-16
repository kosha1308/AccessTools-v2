# Runtime Core Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking. Execution: native, in this chat; independent whole-branch review before publication.

**Goal:** Execute supported typed rules on explicit in-memory data and return typed results and bounded explanations.

**Architecture:** Interpret existing typed AST, with static dispatch by PrimitiveId and selected signature. Separate immutable values/context, traversal/budgets, normalizers and primitive handlers. No source I/O or dynamic code.

**Tech Stack:** VBA, existing Excel COM test runner, PowerShell developer tooling.

**Spec:** `docs/superpowers/specs/2026-10-16-runtime-core-design.md` (approved in chat 2026-10-16).

## Global Constraints

- Offline VBA; no input workbook changes, formulas, arbitrary code or new production dependencies.
- Supported types: TEXT, NUMBER, BOOL, DATE and homogeneous LIST of these types; unsupported domain signatures fail explicitly.
- Locale INVARIANT; comparison BINARY or ASCII_TRIM_CI. Run date and timezone identity supplied explicitly.
- MaxSteps=200000, MaxDepth=64, MaxListItems=10000, MaxTextLength=65536, MaxTraceEntries=20000; depth>128 invalid.
- Existing 953 compiler assertions must remain passing. Compiler catalog is not reduced.
- Native implementation in isolated branch `feat/runtime-core` based on local compiler head; do not merge PR #2.
- Every task: RED evidence, minimal implementation, full GREEN evidence, commit. No speculative implementation before plan approval.

## Review Focus

- Caller mutates input/result list after insertion: active evaluation remains unchanged (task 2).
- Reused evaluator after ERROR or exhausted budget: next run starts with fresh state (task 4).
- Numeric underflow and rounding boundaries: never silently zero or overflow (tasks 3/5).
- Same primitive name with unsupported domain signature: reject before evaluating operands (task 4).
- Trace capacity exhausted while propagating a prior error: preserve primary error, append limit diagnostic; without prior error return RUNTIME_LIMIT_TRACE (task 4).

## Files and shared conventions

All new classes live in `src/runtime/`; handlers in `src/runtime/` standard modules.
Test modules live in `tests/vba/` and are invoked by `modATTestMain.AT_RunTests`.
VBA component names remain unique and at most 31 characters. Public values/results
use private fields and defensive getters, not public mutable Collections.
Runtime error factories use ATDiagnostic with Category=RUNTIME, Severity=ERROR,
and originating node SourceId/Start/Length. Diagnostic IDs are stable ASCII codes.

### Task 1: Complete compiler provenance

**Files:** modify `src/compiler/ATCompiler.cls`, `tests/vba/modATDiagnosticTests.bas`.
**Interfaces:** consumes existing CheckNode/CheckNamed/CheckLookup and ATAstNode.SourceId;
produces complete provenance for existing AST and warning contracts, no API change.

- [ ] Add `NestedWarningSource`, `LookupLiteralSource`, `BindingSource`, `ContextualListSource` regression cases: nested deprecated definition has NAMED:<id> warning; schema-bound source/column literals, BIND and synthetic LIST inherit their original source. Assert exact existing Start/Length as well as SourceId, including DEFAULT:<id>:<parameter> context.
- [ ] Run `powershell.exe -NoProfile -ExecutionPolicy Bypass -File tools/Test.ps1`; expect new provenance assertions FAIL.
- [ ] Set provenance when these nodes are created/bound; set warnings from current definition source, retaining caller source on return.
- [ ] Run full runner; require Compile PASS, 0 FAIL and unchanged catalog; commit `fix: preserve compiler provenance for runtime traces`.

### Task 2: Values, results and explicit context

**Files:** create `ATRuntimeValue.cls`, `ATEvalResult.cls`, `ATEvaluationContext.cls`, `ATTraceEntry.cls`, `modATRuntimeValues.bas` under src/runtime; create `tests/vba/modATRuntimeModelTests.bas`; update test main.
**Interfaces:**
- ATRuntimeValue.InitializeScalar(typeName As String, state As ATValueState, payload As Variant); InitializeList(elementType As String, items As Collection); Clone() As ATRuntimeValue. Read-only DataType/ValueState/Scalar/Count; Item(index As Long) As ATRuntimeValue returns clone. Scalar only PRESENT supported scalar, Item only LIST, else RUNTIME_VALUE_ACCESS. Initialization once only; no implicit string-to-bool/number coercion.
- ATEvalResult.Initialize(state As ATEvalState, value As ATRuntimeValue, diagnostics As Collection, trace As Collection). Read-only State/Value/Diagnostics/Trace; Value getter rejects non-OK. Copy mutable inputs/outputs. INVALID values may carry normalization diagnostics with OK envelope for explicit inspectors.
- ATTraceEntry.Initialize(sourceId As String, start As Long, length As Long, primitiveId As String, status As String, valueState As String, reason As String, ordinal As Long); corresponding read-only properties. Status NOT_EVALUATED distinct from EvalState.
- ATEvaluationContext.Initialize(runId As String, runDate As String, timezoneId As String, localeId As String, comparisonMode As String); AddField(id As String, result As ATEvalResult); Snapshot() As ATEvaluationContext; ResolveField(id As String) As ATEvalResult. IDs canonicalized; duplicate fields rejected. Snapshot frozen, normal context permits adding fields until snapshot.
- Factories `AT_RuntimeScalar(typeName As String, payload As Variant) As ATRuntimeValue`, `AT_RuntimeState(typeName As String, state As ATValueState) As ATRuntimeValue`, `AT_RuntimeOK(value As ATRuntimeValue) As ATEvalResult`, `AT_RuntimeFailure(state As ATEvalState, code As String, node As ATAstNode) As ATEvalResult` in modATRuntimeValues.

- [ ] Write model tests: PRESENT scalar round trip; EMPTY/NULL/MISSING/INVALID remain distinct; blocked payload access; non-OK Value access; wrong native scalar type rejected; duplicate field error; caller/result list mutation cannot alter stored values; snapshot isolates later context additions.
- [ ] Run runner; expect missing classes compile failure.
- [ ] Implement contracts using copies; missing field returns ERROR/RUNTIME_FIELD_MISSING, not a fabricated EMPTY. Invalid config returns stable CONFIG_CONTEXT_REQUIRED/CONFIG_UNSUPPORTED at evaluator boundary.
- [ ] Run full runner, require 0 FAIL; commit `feat: add immutable runtime values and evaluation context`.

### Task 3: Scalar normalization

**Files:** create `modATRuntimeNormalize.bas`, `tests/vba/modATNormalizeTests.bas`; register suite.
**Interfaces:** `AT_NormalizeScalar(typeName As String, raw As Variant, node As ATAstNode) As ATEvalResult`; `AT_ParseDecimal(text As String, node As ATAstNode) As ATEvalResult`; `AT_TextKey(text As String, comparisonMode As String) As String`. Normalization rejection returns OK envelope + INVALID value + diagnostic; evaluator ordinary operations promote unusable value to ERROR. Unsupported type is ERROR/RUNTIME_UNSUPPORTED_NORMALIZER.

- [ ] Add exact assertions: invariant `1.25` ->1.25; `1,25`, suffix garbage, NaN/Infinity invalid; 400-digit magnitude overflow; nonzero decimal below Double range underflow; all-zero tiny decimal remains zero; BOOL True accepted, `"TRUE"` rejected; ISO leap date accepted, 1900-02-29 rejected, years 0100/9999 accepted. Empty TEXT `""` resolves EMPTY while U+0020 text is PRESENT before comparison policy.
- [ ] Run runner; expect missing normalizer compile failure.
- [ ] Implement strict invariant decimal conversion with range guards; reuse compiler Gregorian validation without host locale conversion. Native DATE payload is date-only VBA Date; reject time fraction. ASCII-only case folding and U+0020 trim; preserve Cyrillic and UTF-16 surrogate pairs. Do not change host regional settings.
- [ ] Run cases in both BINARY/ASCII_TRIM_CI profiles and full runner; require 0 FAIL; commit `feat: normalize scalar runtime values explicitly`.

### Task 4: Bounded traversal, logic, explain and named calls

**Files:** create `ATEvaluator.cls`, `ATEvalLimits.cls`, `ATEvalBudget.cls`, `modATRuntimeDispatch.bas`, `tests/vba/modATEvaluatorTests.bas`, `tests/vba/modATEvalLimitTests.bas`; register suites.
**Interfaces:**
- ATEvaluator.Evaluate(compiled As ATCompileResult, context As ATEvaluationContext, Optional limits As ATEvalLimits = Nothing) As ATEvalResult.
- ATEvalLimits exposes spec limit Long properties and Validate().
- ATEvalBudget.Initialize(limits As ATEvalLimits); Step(Optional units As Long=1); EnterNode()/LeaveNode(); CheckList(count As Long); CheckText(length As Long); AddTrace(entry As ATTraceEntry); read-only cloned Trace. Errors RUNTIME_LIMIT_STEPS/DEPTH/LIST/TEXT/TRACE.
- `AT_RuntimeSupports(node As ATAstNode) As Boolean` checks signature and PrimitiveId; `AT_RuntimeVersionSupported(dslVersion As String, schemaVersion As String) As Boolean` accepts only 1.0.0/1.0.0; `AT_ApplyScalar(node As ATAstNode, args As Collection, context As ATEvaluationContext, budget As ATEvalBudget) As ATEvalResult` handler completed task 5. Temporary handler returns unsupported for other primitives until task 5.
- Internal EvaluateNode(node, parameterFrame, depth) handles NUMBER/STRING/BOOL/FIELD/LIST/NAMED and operations; logic/state inspectors handled lazily in evaluator. No mutation of AST and no reparse. Field type checked against node.DataType. Named parameter frames lexical and restored on all exits.

- [ ] Add RED cases: TRUE/FALSE literals; field type mismatch; invalid compilation/absent root -> RUNTIME_COMPILE_REQUIRED; incompatible versions -> RUNTIME_VERSION_UNSUPPORTED (validate version-check helper with supplied version strings since existing getters are constants); unknown Kind -> RUNTIME_UNSUPPORTED_NODE.
- [ ] Add full AND/OR truth/state matrices (OK TRUE/FALSE, ERROR, SKIPPED, MANUAL_REVIEW), NOT, lazy COALESCE; skipped divide-by-zero branch yields NOT_EVALUATED with source/span and no error. State inspector 6x5 matrix exactly matches semantics.md; all non-OK envelopes propagate through each inspector unchanged.
- [ ] Add named default prior-parameter references, supplied-argument caller scope, nested/repeated calls with distinct parameter values, caller parameter-name collision, failure followed by clean run. Assert no frame leak and AST unchanged.
- [ ] Add limits at boundary/boundary+1 for each limit, invalid depth129/zero values, equal input repeatability, trace error precedence, and ERROR then successful run on same evaluator. Check unsupported MATCHES/LIKE/source/row/domain signatures before operand read using deliberately absent field input.
- [ ] Run runner; expect missing evaluator compile failure, then retain failing behavior assertions while implementing traversal, support table, frames, budgets and trace. Short circuit logs one skipped subtree root, not all descendants. Per-element comparison loops consume Step as well as visited nodes.
- [ ] Run full runner; require 0 FAIL; commit `feat: evaluate bounded logic and named expressions with traces`.

### Task 5: Supported scalar/list primitives

**Files:** create `modATRuntimeScalar.bas`, `modATRuntimeSets.bas`, `tests/vba/modATPrimitiveEvalTests.bas`; update dispatch and test main.
**Interfaces:** `AT_ApplyScalar` delegates static groups to `AT_ApplySet(node As ATAstNode, args As Collection, context As ATEvaluationContext, budget As ATEvalBudget) As ATEvalResult`; no dynamic Application.Run. List equality ordered; set comparison ignores duplicates. Scalar payloads are read only after PRESENT/type validation.

- [ ] Add data-driven tests for every supported spec binding: arithmetic/unary, comparisons NUMBER/DATE/TEXT, equality BOOL/LIST, inclusive BETWEEN, text functions, ABS/ROUND, scalar/list MIN/MAX, SUM/AVG, DATE/TODAY/DATE_DIFF, list IN and six set operations. Assert values AND type/state. Positive, invalid-state and boundary fixture for each group.
- [ ] Pin `5.5 % 2 = 1.5`, `-5.5 % 2 = -1.5`, divisor0 -> RUNTIME_DIVISION_BY_ZERO; ROUND(2.5,0)=2, ROUND(3.5,0)=4, precision -1/29/fraction -> RUNTIME_ARGUMENT_RANGE; overflow -> RUNTIME_NUMBER_RANGE. Avoid overflow in intermediate ROUND scaling when final result is representable.
- [ ] Pin MIN/MAX/AVG(empty) -> AGGREGATE_EMPTY; SUM(empty)=0; typed empty lists supplied as fields. Exact Double equality without epsilon. LIST [1,2] != [2,1], EXACT([1,1],[1]) true; empty set laws; invalid list element errors. Charge budget inside nested set comparisons and aggregate scans.
- [ ] Pin LOWER/UPPER ASCII-only, TRIM only external U+0020, LEN supplementary character=2, BINARY vs ASCII_TRIM_CI literal/field comparison symmetry; raw text functions retain text payload policy. TODAY on two explicit run dates differs without system clock calls; DATE_DIFF across leap day signed correctly.
- [ ] Run runner; expect unsupported binding failures. Implement static handlers, normalizers and support signatures together; do not catch expected result objects as native exceptions.
- [ ] Run all suites and unchanged catalog gate; require 0 FAIL; commit `feat: implement scalar and typed-list runtime primitives`.

### Task 6: Acceptance and publication

**Files:** create `tests/vba/modATRuntimeAcceptance.bas`, `docs/runtime-core-status.md`; update test main, README, semantics.md, progress.md. Preserve explicit unsupported matrix for 2B.
**Interfaces:** end-to-end Compile -> Evaluate against immutable synthetic contexts, existing runner unchanged.

- [ ] Add representative full-rule fixtures covering scalar thresholds, named predicates/defaults, list coverage, NULL fallback, INVALID inspection and missing-field failure. Assert deterministic trace and source/span, no source data payload in explain, unchanged AST serialization/context between runs. Compile all deferred operation fixtures and assert explicit runtime unsupported.
- [ ] Run full runner, tooling `tests/tooling/StaticValidate.Tests.ps1` and `tests/tooling/ExcelGate.Tests.ps1`; record observed totals, not target counts. Retain baseline compiler suite plus new assertions.
- [ ] Independent code review against spec and Review Focus; fix important findings with reproducing tests. Record remaining minor findings/rulings in progress.
- [ ] Commit implementation; fresh clone without hardlinks, repeat Excel compile/full tests + static fixtures. If Excel unavailable, record blocked acceptance; static validation alone is insufficient. Do not alter Trust Center.
- [ ] Publish feature branch and draft PR against compiler branch until PR #2 integration is authorized; attach PR. Use available connector or request only required network permission for git, never bypass permission rejection. Verify latest-head GitHub checks before claiming published acceptance.
- [ ] Update status evidence and state next boundary: stage2B domains/rows/sources; no production switch or policy default selected.

## Self-review and execution handoff

Coverage: spec models/task2, normalization/task3, logic/trace/names/budgets/task4,
primitives/task5, provenance/task1, full acceptance/task6. All five Review Focus
conditions have owning tests. ATEvalResult.Value is guarded everywhere; INVALID
inspection is distinguished from non-OK propagation. BIND is provenance-only in
2A because row quantifiers remain unsupported. Native exceptions fail closed.

Clarifications made for implementation: trace exhaustion preserves an already
encountered primary failure and adds a secondary limit diagnostic; empty TEXT is
EMPTY, whitespace normalization does not silently change its ValueState; typed
empty lists come from context because compiler rejects untyped empty literals.

Await plan review before product-code edits. Execution method remains native in
this chat; no further method choice is needed.

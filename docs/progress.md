# Stage 1 progress

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

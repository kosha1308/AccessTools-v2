# PR-02 Schema / XML / Validator Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans inline. Steps use checkbox tracking. The owner's PR-02 task authorizes implementation of this bounded foundation; no PR-03 or runtime checks.

**Goal:** Implement small generic VBA definitions, secure XML persistence and structural diagnostics.
**Architecture:** Two record classes, three production modules, one VBA test module. A kind discriminator and per-kind schema avoid ten near-identical classes. Expressions are deferred structural contracts, operations are externally supplied trusted metadata; nothing evaluates either.
**Tech Stack:** Windows Excel VBA7, late-bound MSXML6 and Scripting.Dictionary/FSO; PowerShell only for developer tests. No third-party runtime files.
**Spec:** User PR-02 task; docs/rule-model.md, expression-model.md and decisions.md at f735d63.

## Global constraints

- No business checks, AST evaluator/type checker, UI, Writer, Rulebook migration or publication controller.
- Unknown XML/schema versions and unsafe XML fail with structured diagnostics.
- Save new revisions only: no overwrite; staging is verified before rename in the same folder. Previous revision is untouched.
- No Trust Center/registry changes; tests own a hidden empty Excel process and temporary files.
- Actual x64 Excel evidence and x86/x64 MSXML probes are distinct from untested x86 VBA.

## Review focus

- External entities/DTD and oversized/deep XML must not be resolved.
- Duplicate properties/references and unknown elements must not be silently dropped.
- Ownership references are not execution dependency edges: only explicit dependency cycles are errors.
- Unicode/whitespace, XML escaping and exact reference versions must survive round-trip.
- Failed save and competing destination must leave previous bytes intact.

## Task 1: runtime probe, tests and models

Files: tests/Run-Tests.ps1, tests/AtSchemaTests.bas, src/CAtDefinition.cls, src/CAtPackage.cls, src/modAtModel.bas.
Interfaces: AtNewPackage(id,name), AtNewDefinition(kind,id,name), AtAddReference(def,role,kind,id,version,expectedType); generic records and diagnostics.
- [x] Write fixtures/tests for two templates, stable IDs and invalid references before production code; observe RED.
- [x] Probe secure MSXML6 in x86/x64 and own Excel/VBProject without changing permissions.
- [x] Implement common records/factories, not template-specific classes.

## Task 2: structural validator

File: src/modAtValidation.bas. Interface AtValidate(package,operationCatalog) As Collection; AtHasErrors(diagnostics).
- [x] Write validator tests before production code, then implement ID/version/type/property/reference/locator checks and explicit-dependency cycle detection.
- [x] Exercise duplicate ID, wrong type, bad/missing ref, foreign ownership, cycle, unknown operation/version and Published metadata failures.
- [x] Do not inspect or execute AST bodies; return a deferred-AST warning for structural-only success.

## Task 3: XML and safe new-revision save

File: src/modAtXml.bas; tests/fixtures/*.xml; docs/adr/008-schema-storage.md.
Interfaces: AtToXml(package), AtFromXml(text,catalog,diagnostics), AtLoad(path,catalog,diagnostics), AtSaveNew(package,path,catalog,diagnostics).
- [x] Write load/round-trip/Unicode/corrupt/unknown-version/DTD/failed-write tests, observe RED.
- [x] Implement bounded closed-shape XML codec and verified staging + non-overwriting rename; no publication/rollback.
- [x] Execute full suite on real available Excel; preserve exact host/bitness and untested limitations in docs/pr02-validation.md.

## Task 4: review and PR

- [x] Review complete diff for genericity, XML trust boundaries, diagnostics and save-failure preservation.
- [x] Run full suite, formatting checks, commit and create PR without merge.

## Execution notes

MSXML6 activation/secure parsing succeeded in x86 and x64 Windows hosts. Initial sandbox Excel activation denied; an authorized isolated hidden Excel probe outside sandbox succeeded (Excel 16.0, VBProject access available). No security settings changed. A task-local clone and dedicated feature branch isolate these changes; no extra worktree was created.

Final evidence: Excel 16.0 build 17932 Office64 PASS|77|0; x86/x64 MSXML probes PASS. Reviewer findings reproduced RED, fixed GREEN and reviewed again. No x86 Excel or UI acceptance claimed. ADR and docs/pr02-validation.md record limits. Final commit/PR link is recorded in the PR report rather than embedded self-referentially here.

# PR-08 implementation plan

> Implementation: execute inline in the existing task; no delegation except read-only review. User spec is attachment037d4cd3, base c4f45ad, separate stacked draft PR; PR39 must not change.

## Contracts and choices

Fields and rules have independent immutable generated IDs. Each rule currently edits one FieldRef, while the underlying AST/reference model supports more fields later. Shared fields are not copied. One immutable snapshot and one union of all active-rule actions per full run; any incomplete required computation freezes writing/cleanup. Selected-rule testing is preview only.

Schema4 adds explicit Rule.Enabled, preserving meanings of1/2/3. EQUAL uses existing operation; NOT_EQUAL adds Not. Requiredness uses a trusted pure Text→Bool operation explicitly consuming only observed EMPTY input obligations; technical errors and StateIs remain unchanged. Whitespace-only Text is empty for requiredness (ASCII Trim), not for equality. New schemas/operations are general, not NETWORK_ACCESS.

Keep legacy CAtAdminEditor API/tests; current launch uses CAtTemplateEditor. Existing form/controller supports the common editor properties and adds field/rule navigation. New runtime findings retain per-rule typed outcomes; initiator copy contains only business messages. Writer remains unchanged, including32intents/targets and512-character Note budgets. Rule copies get new Rule/Expression/Action IDs; fields remain shared.

## Tasks

- [x] Verify682 baseline, v1 control refs and relevant docs/TZ; record only deltas.
- [x] Red/green tests: schema4, trusted requiredness and non-EMPTY error preservation.
- [x] Authoring records and package compile/load/save; stable IDs, enabled/copy/delete/rebind.
- [x] Full-run aggregation and journal lifecycle using unchanged Planner/Writer.
- [x] Current UserForm multi-navigation, three real conditions, all/selected test, multi-results, explicit Writer refusal.
- [x] Synthetic demo≥2fields/3rules and witnesses for all8requested scenarios.
- [x] Full real Excel x64 regression, persisted demo checks, separate nativeUI/manual/DPI limits.
- [x] Read-only review, concise docs/report, commit/push new branch; create draft PR against feat/pr-07-admin-vertical-slice, attach, no merge.

## Files

Schema/runtime changes are limited to existing XML/validator/registry/evaluator/result effect handling. New authoring field/rule/editor classes plus small model compiler and full-run records; existing UI/JSON/build/test tools extended. No adapters, tables, IP/business policies, publication or roles. All Excel tests use synthetic books. Reported r5 red-cell incident has no recoverable output; do not invent its cause. User cannot recall rule/Writer statuses.

# Network user workflow improvements

> Execute inline using superpowers:executing-plans. User requests remembered sources, popup copying all user errors, and cell notes in the request. Existing check semantics remain unchanged.

**Goal:** restore convenient coordinator workflow in AccessToolsv2.xlam; pilot.3.

**Design:** persist only Rulebook/Superset paths in per-user LOCALAPPDATA/AccessToolsv2, separate from immutable package. First selection/missing files or explicit Sources command prompts; each run still re-reads data through unchanged controller. After success prepare routed user errors, show native Excel popup with Copy/Report/Coordinator/Save and allow reopening via Ribbon. Native Unicode clipboard, no external processes. Apply legacy Excel Notes to current request in memory, preserve cell values and foreign note text, remove only owned markers on repeat; never save request automatically.

**Routing:** copying includes only InUserReport + UserError + ERROR. Request notes honor explicit Rulebook routing/highlight; unconfigured input errors get default notes. Additive decision metadata distinguishes explicit settings from absent configuration without changing engine issues/outcomes/routing values. Coordinator-only findings never become user errors.

## Tasks

- [ ] RED/GREEN settings tests: persist/reload Unicode paths, same-path rejection, missing/malformed data, explicit test directory prevents touching user's settings.
- [ ] RED/GREEN feedback tests: exact user-error selection/refs, explicit routing exclusion, merged/foreign cells, user note preservation/repeat/clear, worksheet values unchanged; protected/oversize failure surfaced without breaking results.
- [ ] Thin command integration: retained paths + explicit change, result popup and reopen, copy feedback; no stale result after cancel/failure, no silent copy truncation, menu actions qualified by ThisWorkbook.Name.
- [ ] Rebuild pilot.3 under AccessToolsv2.xlam with updated guide/Ribbon, full regression/tooling, independent review, installed and extracted ZIP smoke. No automatic installation or security changes.

## Acceptance limits

Popup structure tested programmatically without clicking; clipboard deliberately not overwritten by automated tests. Manual popup/copy acceptance on user's Excel remains required. Notes modify only open request metadata and leave workbook unsaved; failures reported visibly. Source settings persistence tests use private temp folders; no real settings paths written by tests.

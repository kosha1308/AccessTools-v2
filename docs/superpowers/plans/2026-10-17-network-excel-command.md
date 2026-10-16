# Network Excel command implementation plan

> Execute inline using superpowers:executing-plans. User delegates implementation choices and asks to continue the agreed workflow.

**Goal:** Russian command invokes the existing controller and saves separate results without overwriting files.

**Architecture:** Thin interactive shell captures active request worksheet before dialogs. Controller stays unchanged. Dedicated exporter writes copies into a newly owned unique subfolder and rolls back only its own partial files on failure. Outputs remain open for review/retry. No startup loading or Ribbon installation in this stage; XLAM packaging next.

**Tech stack:** VBA/Excel COM; supplied XLSX profiles.

## Constraints

Preserve 16 checks and controller contracts. No saved user settings, no input saving, no existing output overwrite. Cancellation is normal. Report business failures separately from run failures. Initial command for coordinator workflow, not final profile-aware release.

## Tasks

- [x] TDD exporter API `ATNetworkOutputSaver.Save(reportBook, coordinatorBook, parentFolder) As String`: output copies with stable Russian names, new timestamped collision-safe folder, only local existing parent, distinct XLSX books, events restored. Test real readable copies, same-parent repeated exports, no mutation of open books/user marker, invalid inputs and partial failure cleanup.
- [x] Russian macro `ПроверитьСвязанности`: capture worksheet; select XLSX paths; capture local Windows date and current UTC offset; display complete vs checks passing vs technical failure; offer folder selection for export, keep unsaved books on cancellation/failure. Guard reentrant calls; restore status bar. Separate `СохранитьРезультатыСвязанностей` macro for current successful result, never stale after failed/cancelled new run.
- [x] Full Excel regression, static tooling, fresh read-only review. Interactive dialog path manually probed if available; explicitly document untested UI branches.
- [x] Document macro invocation and source constraints; clean clone test, push draft PR against PR25. Do not merge/release.

## Review focus

Cancellation before run must not export previous result. Save failure must preserve open outputs. Userbooks never closed. Directory collisions must not touch existing files. Success notification must not claim all business checks passed. Local date and current Windows timezone offset captured explicitly, no guessed zone or date shifted by UTC midnight.

Verification: full/clean clone11303/0, compilePASS, static7/0, gate9/0. Fresh review no blockers. Ruling: partial disk failure/race cleanup reviewed statically; no injected filesystem fault or dialog automation in this stage, explicitly documented as acceptance limits.

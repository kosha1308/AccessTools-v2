# Network run controller implementation plan

Base8e627c6/PR24, inline dedicatedclone/newbranch. Reuse all16checks and writers.
ATNetworkRunController.Run(requestWorksheet,rulebookPath,supersetPath,runId,date,
timezone) -> immutable ATNetworkRunResult with State/Phase/ErrorCode,PolicyReport,
ReportData,ReportBook,CoordinatorBook,sourceFingerprints,RequestUnsaved,IsComplete.
Run first snapshots/gates current request (unsaved edits included), then normalizes
sourcepaths, rejects samefile/requestsource and alreadyopen sourceworkbooks, stamps
path/size/mtime, opens own sourcebooks readonly UpdateLinks0/ForceDisableMacros,
reads actual schemas, compiles policies/index freshly (no stale cache), closes own
books, verifies stamps, runs old14+registry15+Supersetrules16+technicalnotes with
same frozen objects, verifies stamps again and creates separate report/coordinator.
Firstsheet Superset profile; Rulebook7 namedtabs. Local XLSX source profile only.
Existing sourcebook is never saved/closed; caller must save/close before path-run.
StateOK/IsComplete mean pipeline completed, not business checks allPASS; expose
AllChecksPassed separately. No automatic output save, UI/filepick/package next.

On source/gate/system failure, return StateError and actionable blocked report;
no coordinator or staleprior outputs. Retain computed policy report in result if
available. Failed outputs closed before fallback report. Always close own inputs,
restore EnableEvents/AutomationSecurity; never changeTrustCenter/registry.
ATRunSourceStamp(path) immutable fingerprint/VerifyCurrent prevents replaced/deleted
files during run. Stamp freshnessminimum size/mtime; eachnewrun re-reads source even
if fingerprint same. No persistent/sessioncache or hidden compileddata fallback.

TDD: temp artificial rulebook7tabs+Superset and request, atomic16+bothbooks including
bothinstructions, sourceunchanged/ownedclosure/metadata; mutate/replacesamefilename
nextRunnewvalues/no staleafterbadheader; unsavedopen source rejectedwithoutmutation,
unsavedrequestallowed; badgatebeforepaths; pathsamefile/bad/missing; stampchanged;
restoredapplicationstate, failedcleanup, repeat aftererror. FullExcel90s (if new
integration time exceeds envlimit use targeted timings evidence, don'tweakenchecks),
static7/gate9, actualReadOnlyreference, freshreview, cleanclone,draftPRbase24.

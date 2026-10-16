# Coordinator technical notes implementation plan

Baseb87f2bb/PR23. Inline, dedicatedclone/newbranch. Scope is bounded adapter over
finished checks, no re-evaluation or parsing of issue message semantics.
ATCoordinatorTechnicalNotes.Run(source,registry,policyReport,optionalBudget)
returns immutable ATCoordinatorNotesResult(State,Rows,GlobalCount,Incomplete)
bound to registry. Only policyReport.SupersetIssues for NET.IP_EXCEPTIONS are
copied, preserving coordinator message/routing applied once. Require exactlyone
technicaloutcome; blocked/system/budget -> emptyRows/error, no partial notes.
Associate issue.CellRef to actual RAW technical row (source or dest refs). Attach
full connection note to entries directly referencing that row endpoint or
mathematically overlapping valid whole-cell endpoint domain. Groups/invalid
cells not salvaged; unassigned/global issues counted and remain in existing
check report. No arbitrary rule scope/port guessing, no service attribution:
connection-scoped technical service/output stays inside labelled instruction.

Writer.Create(registry,optionalSupersetAnnotations,optionalTechnicalNotes):
validate identity/rowids beforecreating; append F, preserve E and5infofields,
error fallback explicitly says notes nottransferred and retains basebook;
conflict/error highlight wins; banners report incomplete/global notes and point
to separate issue report. No17thcheck, no rerun engine, no sourcefilewrites/save.
Public old1/2arg calls compatible. Source/policy report must come from same
controller run; sourcefilefreshness and atomic orchestration remain nextstage.

TDD: routed technical rule port443 context, source/dest+lower subset vs unrelated,
multirulemessages+refs preserved, no extra user errors or outcomechanges; policy
customcoordinatormessage once, disabled routeexcluded, globalmanual visible,
identity/errors/budget no partial; writer append notoverwrite and failfallback.
FullExcel90s, static7/gate9, ReadOnlyreferences, freshreview, cleanclone, demo,
draftPR basePR23 no merge/release. Runtime200000steps,10000notes,32767celltext.

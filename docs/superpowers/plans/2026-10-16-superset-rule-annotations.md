# Superset rule annotations Implementation Plan

Base342e4f2/PR22. Inline, dedicatedclone/newbranch. Existing engine/checks unchanged.
Goal: evaluate frozen Rulebook Superset contexts, annotate coordinator workbook,
NET.HOSTNAME_RULES as independent16thcheck; no fabricated ports or arbitrary host.

1. ATNetworkSupersetRules.Run(registry,rules,optionalBudget) -> frozen
ATNetworkSupersetResult(State,Annotations,Report,sourceRegistry identity guard).
Annotations COLLECTION<ROW<SUPERSET_ANNOTATION>> bysummaryId, SERVICE_RESULT,
INSTRUCTION,STATUS. Bind to exact registry object to prevent stale/wrong report.
For each entry/matched source row, intersect entry domain with valid whole-cell
source IP domains; invoke unchanged Policy.Superset(intersectionText,rawhostname,
rawservice,sharedbudget). Invalid sourceIP cells never salvaged, valid othercells
remain. Intersection empty blocks invalid binding. Networks not expanded.
No matches -> empty annotations + explicit manual note, no guessed service/host.
2. Existing specificity/ties/conflicts/severity remain engine-owned. Confirmed
contexts retained when other contexts are manual; system/budget error discards
new partial annotations, reportBLOCKED. Perentry services unique multiline;
full rule outputs/instructions/source refs retained in instruction blocks, same
rule outputs dedup perentry but issue contexts identify actual Superset row.
Different valid services from different source rows are not invented conflicts.
990issues/32767text/10000entries/shared200000 budget; source incomplete ->SKIP.
3. Writer.Create(registry, optionalAnnotations) preserves oldoneargbehavior.
With valid bound annotations fills E/F and explains Superset rules scope.
Blocked annotations retain base book and explicit notchecked instructions;
wrong registry identity rejects beforebookcreate. Doesnot alter raw8columns or
fivefieldsenrichment. Rule/service applicability uses original source service,
never another rule's output (no chaining).
4. ApplySupersetRules(previous15,annotations,catalog) adds16 independently,
keeps15 and decisions on new failure; publishedkey NET.HOSTNAME_RULES|severity|
Имя хоста; conflict key Конфликт Rulebook; coordinator/Superset fallback, no
userinputerror. UI titles/banner expose actualscope without claiming release.
5. TDD: hostname specificity + original service conditions, noresultchaining,
network/row IP intersection falsematch guard, multiple services nonconflicting,
ties conflict, missing data/manual/invalidsource, bounds/no partial, writer
E/F+identityguard+basefallback,16preserves15/severity/templateoverflow.
Full Excel90s, referencesReadOnly, static7/gate9, onefreshreview, cleanclone,
syntheticexample/visualQA, draftPR basePR22, nomerge/release.

Next: technical-rule annotations correlation/controller freshness+unsaved edits,
user commands/save packaging/full scenario acceptance. This stage does not
invent source/destination roles or ports from an IP-only Superset context.

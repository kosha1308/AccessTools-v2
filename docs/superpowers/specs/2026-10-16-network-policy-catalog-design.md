# Validated Rulebook catalog and local policy application

Continue from1369b65 without rewriting completed checks. The user delegates
technical execution and requests continued development, not approval loops.
This stage connects four RAW tables: PROTOCOLS, PROBLEMS, ROUTING, COMMON_TEXT.
Initiator/IP/hostname/service rule evaluation and Superset remain subsequent stages.

Compile immutable policy from four successful typed source snapshots. Active flags
must be explicit Да/Нет, YES/NO, TRUE/FALSE or1/0; unknown/blank flags fail rather
than imply permission. Inactive rows do not override behavior. Active duplicate
problem/routing/common keys fail; duplicate canonical protocols deduplicate.
Protocols match the existing checker contract: at most1000 unique entries,
1–128 ASCII characters A–Z,0–9,_-./+ after canonicalization.
Only scoped keys check|original severity|field are supported, no wildcard guessing.
Keys resolve before severity changes. Header-normalized field tags are supported;
check ID and severity components validated. Severity supports Russian/English labels.
Errors in active used cells, absent fields, malformed types or >10000 rows/table
fail closed. Shared compile budget200000 steps including consumed text lengths;
65536 text per value; stored messages32767. Blocked/configuration failures remain
visible but are not classified as user-input errors.

Blank custom user/coordinator messages retain original computed issue text;
coordinator falls back to selected user text. Builtin default-message text is not
reused as a template because it may contain unresolved business placeholders.
Custom templates support proven {message},{field},{check},{severity},{row} tokens;
reject unknown/unbalanced tokens, expand once without interpreting inserted text,
and reject >32767 expanded characters. Common-text values remain data, not templates.

Apply policy through a new wrapper runner, passing compiled active protocols into
ATNetworkLocalChecks unchanged. Explicit mapping translates existing issue codes
and fields into published scoped keys. Unmapped issues retain existing text/severity.
Execution failures and blocked reports bypass presentation policy and cannot hide.
For completed reports, recalculate PASS/FAIL after severity changes while preserving
SKIPPED/BLOCKED. Routing never changes underlying execution status.

Immutable decisions expose configured user/coordinator messages, severity,
UserError and flags for user report/request/Superset. If no custom routing exists,
retain existing user-report inclusion; enable no new request/Superset side effects.
Request routing also requires Problems.HIGHLIGHT. UserErrorCount counts included
ERROR decisions marked UserError; AllReport.ErrorCount still counts all ERRORs.
Keep both explicit; do not mislabel approval as user error. UI consumer must use
routed decisions, not treat AllReport as filtered user report. This stage adds no
source writes or source/request annotations. Unknown future issue keys need explicit
binding before applying policy. Corrupt policy yields one visible NET.POLICY BLOCKED
report with original diagnostics, no partial successes.

Tests use fictional source rows and current form fixture; real schema Rulebook is
read-only acceptance only. No business data/code copied from the legacy repository.

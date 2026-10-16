# Initiator admission, approvals and IP access

Continue base8ad6a2c, preserve completed local checks/catalog. Native inline execution
under existing technical delegation. User explicitly decided16Oct2026: if neither
email nor domain has IP policy, there are no personal IP restrictions. Publish this
as NETWORK NoPolicyAction=ALLOW, while generic compiler requires an explicit action.
This never replaces the independent allowed-initiator check.

Two RAW inputs INITIATORS/INITIATOR_IP compile immutable actor policy. Active identity
rows admit matching selectors except explicit APPROVAL_ONLY/«Только согласование»
service mode; empty service mode is ordinary new-format admission. Other modes fail.
Identity active flags explicit. Matching approval rows accumulate independently;
empty/invalid email skips both identity checks. No active admission rows ->SKIPPED,
not a guessed allowlist. Empty approval set ->PASS. Empty approval severity can use
legacy WARNING; explicit invalid labels block. Scoped catalog may override severity.

Selectors support exact email and exact domain (optional leading@), case-insensitive;
comma/semicolon/pipe/ASCII whitespace lists, exclusions НЕ:/NOT:. A negative selector
excludes that row, never turns into an alternative positive match. Require positive
selectors and strict whole-token email/domain syntax. For IP policy only, matching
positive exact-email rows replace domain rows globally; no domain widening of a
personal policy. Combine selected same-scope/Любой rows by union allow and union deny.
Request fully subset of allow AND no overlap with deny; no partial coverage PASS.

Published XLSX header comments ground IP-table defaults (different from catalogs):
Active empty=Да; scope empty=Любой; scopes Любой/Источник/Назначение;
DB_ACCOUNT empty/Не ограничено=no extra constraint, Запрещено=request account blank.
Any matching selected row forbidding account enforces it at real C10 ref.
Empty allow is the empty set, not an implicit universe; universal allow is0.0.0.0/0.
Allow may contain НЕ:/NOT: terms appended to deny. Explicit deny field accepts IP
expressions, not double-negative instructions. IPv4/CIDR/full range/wildcards x/х
compile through existing interval math; no core modifications or partial salvage.
Unknown/invalid policy data fails closed. At most10000 rows/table, shared200000steps,
65536chars per value,10000 intervals; text lengths and runtime work consume budget.

Actor check scope NETWORK_INITIATOR has3outcomes: INITIATOR_ALLOWED,
INITIATOR_APPROVAL_RULES, INITIATOR_IP_ACCESS. Check technical SOURCE/DEST separately;
include lower actual IPs. Lower expressions already covered by successfully checked
upper numeric addresses need no repeated directional assignment. Otherwise evaluate
ANY policy, and if selected directional rules exist, explicitly require manual review
rather than guess source/destination. Unsupported/group/invalid input creates explicit
incomplete result; keep checking other cells so denied numeric addresses remain visible.
Per-cell false ->ERROR with offending original text/ref, any violation beats skipped;
no false healthy result after work/system errors. DB restriction shares IP outcome.

Composite runner aggregates raw local10 +actor3 then applies presentation policy once.
Extract existing policy wrapper Apply(report,loaded) without changing Run semantics.
Approval fallback remains visible in request, not user error/report or Superset unless
explicit catalog/routing overrides. AllReport retains it. Add exact bindings/titles.
No source annotations or startup source loading. Not full NETWORK_ACCESS readiness:
IP/hostname/service rule evaluation, Superset matcher, coordinator/UI/package remain.

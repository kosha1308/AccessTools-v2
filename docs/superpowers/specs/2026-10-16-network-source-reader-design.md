# Published source schema and read-only snapshots

User supplied empty Superset and NETWORK_ACCESS Rulebook in the templates folder.
Goal: read their actual columns without source mutation or invented policy.
This is the prerequisite for the already authorized matcher/coordinator stage.
Native inline execution; technical decisions remain delegated by the user.

Reuse ATExcelTableSpec header normalization/location and runtime row snapshots.
Add a schema factory for SUPERSET and seven data sheets of the published Rulebook:
INITIATORS, RULES, PROTOCOLS, PROBLEMS, COMMON_TEXT, ROUTING, INITIATOR_IP.
The contents sheet is documentation, not a rule table. Headers may move within
rows1–20/columns1–64; require one unmerged header row. Match canonical full headers,
reject missing/duplicate/ambiguous headers. Hidden source columns remain readable.

Superset IP inputs are the three SOURCE columns, not the unnamed numeric export
index. Preserve FQDN, service, IS, SKUF and support fields. No IP extraction/matching
in this task. Retain all source rows, including filtered/hidden rows, in the raw
snapshot; final-view enrichment retains its existing visible-row semantics.

ATNetworkSourceReader.Read(ws,sourceId) returns ATEvalResult with typed raw rows.
Find last actual value across mapped fields, ignore formatting and unrelated index.
Read bounded array up to100000 body rows and1000000 rectangular cells; exceedance
is SOURCE_ROW_LIMIT/SOURCE_CELL_LIMIT with no partial value. Source flat tables
cannot contain merged mapped cells. Reject stale schema by locating each run.
Preserve original worksheet row IDs, per-field refs, text/empty and cell errors;
no guessed active flags, default protocols or effective policy. Snapshots survive
source closure. No automatic file opening at startup and no source writes.

Do not commit private XLSX or its business values. Schema names only in public code;
tests use fictional records. Supplied books are inspected read-only with hashes.
User-ready v2 still requires actual matching, policy evaluation, summary and package.

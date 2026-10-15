# Stage 1 — compiler foundation milestone

Implemented in the independent AccessTools-v2 repository:

- Windows developer runner with explicit VBAProject compilation, bounded worker,
  test result validation and cleanup of its own Excel process.
- Static validation and malformed component fixtures; hosted static CI.
- Schema/DSL version contracts, diagnostic spans, immutable primitive registry.
- Lexer, shared AST/parser, canonical textual serializer and round-trip checks.
- Field/scope resolution, signature checking, named/default expansion and cycles,
  source dependency metadata and UsedBy edges.
- Input/token/AST/depth/collection/named/work budgets; seeded malformed-input checks.
- Appendix-N compiler examples with explicit diagnostics for unavailable capabilities.

Last development-tree verification: Compile VBAProject PASS; VBA 787 assertions
PASS / 0 FAIL; static validator fixtures 7 PASS; Excel gate fixtures 6 PASS. These
are compiler/tooling checks, not a claim of production Runtime readiness.

An independent review found five important issues: recursion through defaults,
unknown metadata types, malformed JSON counters, syntax-vs-ID parser lookup and
ambiguous registered syntax. Each has been addressed with regression coverage.
Deprecated named expressions now retain compatibility with a warning; duplicate
parameter IDs fail. A test initially indexing an absent warning crashed its Excel
worker; the test guard was corrected and the subsequent complete run is green.

## Acceptance remains open

Do not begin a real-request evaluator cutover from this status alone. The complete
stage-one DoD still requires review/finalization of the supported grammar and
catalog against all intended capabilities, domain literal/normalizer contracts,
and broader canonical/property/golden coverage (including explicit assertions
on diagnostic codes for every intentionally unsupported appendix example).
The full evaluator registry, actual IP/port algorithms, source adapters and UI are
not present. No AccessTools 0.2.x production check or file has changed.

This is a working, tested compiler foundation, not completion of the whole v0.13 TZ.

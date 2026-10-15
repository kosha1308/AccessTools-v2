# Stage 1 acceptance — compiler profile 1.0.0

Status: compiler-stage acceptance is ready, subject to the clean-checkout and CI
results recorded in docs/progress.md. This is the Schema/DSL/Parser/Validator stage,
not completion of the whole AccessTools-v2 product.

## Delivered and checked

- Independent VBA source, static checks, explicit Compile VBAProject and developer
  runner; no checkout/artifacts of AccessTools 0.2.x are needed.
- Versioned type/primitive/normalizer contracts; immutable registered signatures.
- Shared lexer/AST/parser/serializer, scalar/contextual/list/set/Boolean/arithmetic
  syntax, multiword operators, BETWEEN and registered functions.
- Type/scope checking, typed source-schema LOOKUP/SOURCE_VALUE, dependencies/UsedBy,
  hygienic named definitions/defaults, parsed parameter bindings and cycles.
- Exact canonical decimals, ISO Gregorian date and known port-literal validation.
- Stable diagnostic codes, original spans and named/default source provenance.
- Bounded input/tokens/nodes/depth/collections/named expansion/work; seeded negative
  inputs, round-trip and golden diagnostic/compiler fixtures.
- Registry-generated language catalog; every operator/function example compiles.
  Missing/outdated snapshots fail the Excel gate instead of silently drifting.

Development-tree evidence: Compile VBAProject PASS; 953 VBA assertions expected
after the final fixture run; static tooling and Excel gate integration are run
again before publication. Exact observed totals are recorded in progress.md.
No readiness claim for production Runtime is inferred from compiler-only tests.

Appendix-N examples compile with appropriate typed metadata. DISTINCT_COUNT on a
scalar IP produces TYPE_SIGNATURE_MISMATCH; with an explicit LIST<IP_EXPRESSION>
binding it compiles. Legacy ALLROWS colon input is deliberately rejected with
PARSE_BARE_IDENTIFIER until a separate migration importer exists.

## Next stage boundary

Next work is an isolated Runtime Evaluator implementing EvalResult/ValueState,
the registered normalization/evaluation contracts, and explain traces. Begin with
synthetic contexts; real-request shadow mode follows those tests. No production
0.2.x check has been replaced. No Admin editor, production publisher, release ZIP,
real source adapter or runtime IP/port interval evaluator is included here.

NoPolicyAction for missing email/domain-to-IP policy remains a mandatory explicit
business setting before that policy is published; this compiler does not choose it.

## Review follow-ups

The final independent static review found no important readiness blocker.
A non-compiling NOT_EMPTY documentation example prompted tests for every registry
operator/function example; all placeholder arithmetic/logic examples were corrected.

Two smaller explain-stage follow-ups are retained:
- Nested deprecation-warning SourceId currently defaults to RULE.
- Directly typed source/lookup literals, BIND and synthetic contextual LIST nodes
  can omit AST SourceId. Compilation diagnostics/errors retain their own provenance.

These concern future explain output, not typed rule acceptance or data access.
They must be addressed when implementing the evaluator/explain stage.
